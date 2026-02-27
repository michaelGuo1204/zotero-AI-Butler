import { getPref } from "../utils/prefs";
import { PDFExtractor } from "./pdfExtractor";
import JSZip from "jszip";

/**
 * MineruClient handles the Zotero to MinerU API interaction.
 * It uses the 'let it crash' pattern for parsing the somewhat undocumented batch API response.
 */
export class MineruClient {
    /**
     * Main entry to extract markdown from a Zotero PDF item using MinerU
     */
    public static async extractMarkdown(item: Zotero.Item): Promise<string> {
        const apiKey = (getPref("mineruApiKey") as string) || "";
        if (!apiKey) {
            throw new Error("MinerU API Key not configured.");
        }

        // Get PDF file path
        const pdfAttachments = await PDFExtractor.getAllPdfAttachments(item);
        if (!pdfAttachments || pdfAttachments.length === 0) {
            throw new Error("No PDF attachment found.");
        }
        const pdfAttachment = pdfAttachments[0];
        const filePath = await pdfAttachment.getFilePathAsync();
        if (!filePath) {
            throw new Error("PDF file path not found.");
        }

        ztoolkit.log(`[MineruIntegration] Starting MinerU parsing of ${filePath}`);

        // Read PDF binary
        const fileData = await IOUtils.read(filePath);

        // Get Batch & Upload URLs
        // Assuming simple payload for /api/v4/file-urls/batch based on standard implementations
        const fileName = "document.pdf";
        const batchRes = await fetch("https://mineru.net/api/v4/file-urls/batch", {
            method: "POST",
            headers: {
                Authorization: `Bearer ${apiKey}`,
                "Content-Type": "application/json",
            },
            body: JSON.stringify({
                files: [{ name: fileName }],
            }),
        });

        if (!batchRes.ok) {
            const err = await batchRes.text();
            throw new Error(`Failed to get upload URL: ${err}`);
        }

        const batchData = (await batchRes.json()) as any;
        let putUrl = "";
        const batchId = batchData?.data?.batch_id;

        // Dynamic property search, crash if not found
        if (batchData?.data?.file_urls?.[0]) putUrl = batchData.data.file_urls[0];
        else if (batchData?.data?.urls?.[0]) putUrl = batchData.data.urls[0];
        else if (batchData?.data?.urls?.[fileName])
            putUrl = batchData.data.urls[fileName];
        else if (batchData?.data?.items?.[0]?.url)
            putUrl = batchData.data.items[0].url;
        else if (batchData?.data?.upload_url) putUrl = batchData.data.upload_url;

        if (!putUrl || !batchId) {
            throw new Error(
                `MinerU API returned unexpected batch response: ${JSON.stringify(batchData)}`,
            );
        }

        // Upload file content to the presigned URL
        ztoolkit.log(`[MineruIntegration] Uploading PDF to Mineru PUT URL...`);
        const putRes = await fetch(putUrl, {
            method: "PUT",
            body: fileData,
        });

        if (!putRes.ok) {
            const errText = await putRes.text();
            throw new Error(
                `Failed to put upload file, status: ${putRes.status}, error: ${errText}`,
            );
        }

        // Poll for task completion
        ztoolkit.log(
            `[MineruIntegration] Polling for task completion... Batch ID: ${batchId}`,
        );
        return await this.pollStatusAndDownload(apiKey, batchId);
    }

    private static async pollStatusAndDownload(
        apiKey: string,
        batchId: string,
    ): Promise<string> {
        const url = `https://mineru.net/api/v4/extract-results/batch/${batchId}`;

        // Poll for up to 5 minutes (60 * 5s)
        for (let i = 0; i < 60; i++) {
            await Zotero.Promise.delay(5000);

            const res = await fetch(url, {
                headers: {
                    Authorization: `Bearer ${apiKey}`,
                },
            });
            const data = (await res.json()) as any;

            // Batch result usually returns an array under extract_result
            const result = data?.data?.extract_result?.[0] || data?.data;
            const state = result?.state;

            if (state === "done") {
                const zipUrl = result?.full_zip_url;
                if (!zipUrl) {
                    throw new Error(
                        "MinerU Task completed but no full_zip_url returned.",
                    );
                }
                return await this.downloadAndExtractMarkdown(zipUrl);
            } else if (state === "error") {
                throw new Error(`MinerU Task failed processing.`);
            }

            // continue polling...
        }
        throw new Error("MinerU Task timed out.");
    }

    private static async downloadAndExtractMarkdown(
        zipUrl: string,
    ): Promise<string> {
        ztoolkit.log(`[MineruIntegration] Downloading zip result from ${zipUrl}`);
        const res = await fetch(zipUrl);
        if (!res.ok) {
            throw new Error(`Failed to download zip file from ${zipUrl}`);
        }
        const arrayBuffer = await res.arrayBuffer();

        // Extract using JSZip
        // Zotero/Firefox extension environment does not natively provide setImmediate which JSZip needs
        if (typeof (globalThis as any).setImmediate === "undefined") {
            (globalThis as any).setImmediate = (fn: (...args: any[]) => void) =>
                setTimeout(fn, 0);
        }

        const zip = new JSZip();
        await zip.loadAsync(arrayBuffer);

        let mdContent = "";

        // Find the first valid markdown file
        const mdFiles = Object.values(zip.files).filter(
            (file) => file.name.endsWith(".md") && !file.name.includes("__MACOSX"),
        );

        if (mdFiles.length > 0) {
            mdContent = await mdFiles[0].async("string");
        }

        if (!mdContent) {
            throw new Error("No valid Markdown file found in the extracted zip.");
        }

        return mdContent;
    }
}
