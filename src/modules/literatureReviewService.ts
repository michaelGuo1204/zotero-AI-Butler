/**
 * ================================================================
 * 文献综述服务
 * ================================================================
 *
 * 本模块提供文献综述生成的核心服务
 *
 * 主要职责:
 * 1. 创建报告条目
 * 2. 将选中的 PDF 作为附件添加到报告
 * 3. 逐篇文献按表格模板填表（并行，可复用已有表格）
 * 4. 汇总表格内容生成文献综述
 * 5. 生成 AI 笔记并关联到报告条目
 *
 * @module literatureReviewService
 * @author AI-Butler Team
 */

import { PDFExtractor } from "./pdfExtractor";
import { NoteGenerator } from "./noteGenerator";
import LLMClient from "./llmClient";
import { getPref } from "../utils/prefs";
import { ProviderRegistry } from "./llmproviders/ProviderRegistry";
import { PdfFileInfo } from "./llmproviders/ILlmProvider";
import { marked } from "marked";
import {
  DEFAULT_TABLE_TEMPLATE,
  DEFAULT_TABLE_FILL_PROMPT,
  DEFAULT_TABLE_REVIEW_PROMPT,
} from "../utils/prompts";

/** AI-Table 标签名，用于标识文献填表笔记 */
const TABLE_NOTE_TAG = "AI-Table";

/**
 * PDF 文件信息（带文件路径）
 */
interface PdfFileData {
  title: string;
  filePath: string;
  content: string;
  isBase64: boolean;
}

/**
 * 文献综述服务类
 */
export class LiteratureReviewService {
  /**
   * 生成文献综述（表格驱动的两阶段流程）
   *
   * 流程:
   * 1. 创建报告条目
   * 2. 添加 PDF 附件到报告
   * 3. 逐篇填表（并行，复用已有表格）
   * 4. 汇总所有表格 → 调用 LLM 生成综述
   * 5. 创建综述笔记
   *
   * @param collection 目标分类
   * @param pdfAttachments 选中的 PDF 附件
   * @param reviewName 综述名称
   * @param prompt 用户自定义综述提示词（可选，默认使用 tableReviewPrompt）
   * @param progressCallback 进度回调
   * @returns 创建的报告条目
   */
  static async generateReview(
    collection: Zotero.Collection,
    pdfAttachments: Zotero.Item[],
    reviewName: string,
    prompt: string,
    progressCallback?: (message: string, progress: number) => void,
  ): Promise<Zotero.Item> {
    // 1. 逐篇填表阶段
    const tableTemplate =
      (getPref("tableTemplate" as any) as string) || DEFAULT_TABLE_TEMPLATE;
    const fillPrompt =
      (getPref("tableFillPrompt" as any) as string) ||
      DEFAULT_TABLE_FILL_PROMPT;
    const concurrency = (getPref("tableFillConcurrency" as any) as number) || 3;

    // 构建父条目 → PDF 附件的映射
    const itemPdfPairs: Array<{
      parentItem: Zotero.Item;
      pdfAttachment: Zotero.Item;
    }> = [];
    for (const pdfAtt of pdfAttachments) {
      const parentID = pdfAtt.parentID;
      if (parentID) {
        const parentItem = await Zotero.Items.getAsync(parentID);
        if (parentItem) {
          itemPdfPairs.push({ parentItem, pdfAttachment: pdfAtt });
        }
      }
    }

    progressCallback?.("正在逐篇填表...", 10);

    const tableResults = await this.fillTablesInParallel(
      itemPdfPairs,
      tableTemplate,
      fillPrompt,
      concurrency,
      (done, total) => {
        const progress = 10 + Math.floor((done / total) * 50);
        progressCallback?.(`正在填表 (${done}/${total})...`, progress);
      },
    );

    // 2. 汇总表格并生成综述
    progressCallback?.("正在汇总表格...", 65);

    const aggregated = this.aggregateTableContents(tableResults, itemPdfPairs);

    progressCallback?.("正在生成综述...", 70);

    const reviewPrompt =
      prompt ||
      (getPref("tableReviewPrompt" as any) as string) ||
      DEFAULT_TABLE_REVIEW_PROMPT;
    const fullPrompt = `${reviewPrompt}\n\n以下是各文献的结构化信息表格：\n\n${aggregated}`;

    let summaryContent = await LLMClient.generateSummaryWithRetry(
      aggregated,
      false,
      fullPrompt,
    );

    // 3. 后处理引用链接
    summaryContent = await this.postProcessCitations(
      summaryContent,
      itemPdfPairs,
    );

    progressCallback?.("正在创建笔记...", 90);

    // 4. 创建独立笔记（直接放在分类目录下）
    const reviewNote = await this.createStandaloneReviewNote(
      collection,
      reviewName,
      summaryContent,
    );

    // 5. 为所有已纳入综述的文献添加 AI-Reviewed 标签
    for (const { parentItem } of itemPdfPairs) {
      try {
        const existingTags: Array<{ tag: string }> =
          (parentItem as any).getTags?.() || [];
        if (!existingTags.some((t) => t.tag === "AI-Reviewed")) {
          parentItem.addTag("AI-Reviewed");
          await parentItem.saveTx();
        }
      } catch (e) {
        ztoolkit.log(
          `[AI-Butler] 添加 AI-Reviewed 标签失败: ${parentItem.getField("title")}`,
          e,
        );
      }
    }

    progressCallback?.("完成!", 100);

    return reviewNote;
  }

  // ==================== 表格填写相关方法 ====================

  /**
   * 对单篇文献的 PDF 进行填表
   *
   * @param item 文献条目
   * @param pdfAttachment PDF 附件
   * @param tableTemplate Markdown 表格模板
   * @param fillPrompt 填表提示词
   * @param progressCallback 进度回调
   * @returns 填好的 Markdown 表格字符串
   */
  static async fillTableForSinglePDF(
    item: Zotero.Item,
    pdfAttachment: Zotero.Item,
    tableTemplate: string,
    fillPrompt: string,
    progressCallback?: (message: string, progress: number) => void,
  ): Promise<string> {
    const itemTitle = (item.getField("title") as string) || "未知标题";

    progressCallback?.(`正在提取 PDF: ${itemTitle.slice(0, 30)}...`, 10);

    // 提取 PDF 内容
    const filePath = await pdfAttachment.getFilePathAsync();
    if (!filePath) {
      throw new Error(`PDF 附件无文件路径: ${pdfAttachment.id}`);
    }

    let pdfContent: string;
    let isBase64 = false;

    try {
      const fileData = await IOUtils.read(filePath);
      pdfContent = this.arrayBufferToBase64(fileData);
      isBase64 = true;
    } catch (e) {
      // 回退到文本模式
      pdfContent = await PDFExtractor.extractTextFromItem(item);
      isBase64 = false;
    }

    // 构建完整提示词：将 ${tableTemplate} 替换为实际模板
    const actualPrompt = fillPrompt.replace(
      /\$\{tableTemplate\}/g,
      tableTemplate,
    );

    progressCallback?.(`正在填表: ${itemTitle.slice(0, 30)}...`, 50);

    // 调用 LLM 填表
    const result = await LLMClient.generateSummaryWithRetry(
      pdfContent,
      isBase64,
      actualPrompt,
    );

    progressCallback?.(`填表完成: ${itemTitle.slice(0, 30)}`, 100);

    return result;
  }

  /**
   * 查找文献条目是否已有 AI-Table 填表笔记
   *
   * @param item 文献条目
   * @returns 填表笔记内容，未找到返回 null
   */
  static async findTableNote(item: Zotero.Item): Promise<string | null> {
    try {
      const noteIDs = (item as any).getNotes?.() || [];
      for (const nid of noteIDs) {
        const note = await Zotero.Items.getAsync(nid);
        if (!note) continue;
        const tags: Array<{ tag: string }> = (note as any).getTags?.() || [];
        const hasTableTag = tags.some((t) => t.tag === TABLE_NOTE_TAG);
        if (hasTableTag) {
          const noteContent: string = (note as any).getNote?.() || "";
          // 提取 data-ai-table-raw 元素中的原始 Markdown（兼容 div 和 pre）
          const rawMatch = noteContent.match(
            /<(?:div|pre)[^>]*data-ai-table-raw[^>]*>([\s\S]*?)<\/(?:div|pre)>/,
          );
          if (rawMatch && rawMatch[1]) {
            // 反转义 HTML 实体
            const raw = rawMatch[1]
              .replace(/&amp;/g, "&")
              .replace(/&lt;/g, "<")
              .replace(/&gt;/g, ">")
              .replace(/&quot;/g, '"')
              .trim();
            return raw || null;
          }
          // 兼容旧格式：直接去除 HTML 标签
          const textContent = noteContent.replace(/<[^>]*>/g, "").trim();
          return textContent || null;
        }
      }
      return null;
    } catch {
      return null;
    }
  }

  /**
   * 保存填表结果为子笔记（AI-Table 标签）
   *
   * 如果已存在 AI-Table 笔记，则跳过（不覆盖）
   *
   * @param item 文献条目
   * @param tableContent 填表的 Markdown 内容
   * @returns 创建的笔记，或已存在的笔记
   */
  static async saveTableNote(
    item: Zotero.Item,
    tableContent: string,
  ): Promise<Zotero.Item> {
    // 检查是否已有 AI-Table 笔记
    const existingContent = await this.findTableNote(item);
    if (existingContent) {
      // 已存在则跳过，找到并返回已有笔记
      const noteIDs = (item as any).getNotes?.() || [];
      for (const nid of noteIDs) {
        const note = await Zotero.Items.getAsync(nid);
        if (!note) continue;
        const tags: Array<{ tag: string }> = (note as any).getTags?.() || [];
        if (tags.some((t) => t.tag === TABLE_NOTE_TAG)) {
          return note;
        }
      }
    }

    // 创建新的填表笔记
    // 不使用 formatNoteContent，避免标题模式与 AI 笔记冲突
    const itemTitle = ((item.getField("title") as string) || "未知").slice(
      0,
      60,
    );

    // 使用 marked 将 Markdown 表格转换为 HTML 表格（用于 Zotero 显示）
    marked.setOptions({ gfm: true, breaks: true });
    let renderedHtml = marked.parse(tableContent) as string;
    // 移除内联样式，Zotero 笔记不支持
    renderedHtml = renderedHtml.replace(/\s+style="[^"]*"/g, "");

    // 将 LaTeX 公式转换为 Zotero 原生格式
    // 块级公式: $$...$$ → <span class="math">$\displaystyle ...$</span>
    renderedHtml = renderedHtml.replace(
      /\$\$([\s\S]*?)\$\$/g,
      (_match, formula) =>
        `<span class="math">$\\displaystyle ${formula.trim()}$</span>`,
    );
    // 行内公式: $...$ → <span class="math">$...$</span>
    // 使用负向前瞻/后瞻避免匹配已处理的 $$
    renderedHtml = renderedHtml.replace(
      /(?<!\$)\$(?!\$)([^$\n]+?)(?<!\$)\$(?!\$)/g,
      (_match, formula) => `<span class="math">$${formula.trim()}$</span>`,
    );

    // 将原始 Markdown 存储在隐藏元素中，供 findTableNote 提取
    const escapedRaw = tableContent
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
    const noteHtml =
      `<h2>📊 文献表格 - ${itemTitle}</h2>` +
      `<div>${renderedHtml}</div>` +
      `<div style="display:none" data-ai-table-raw>${escapedRaw}</div>`;

    const note = new Zotero.Item("note");
    note.libraryID = item.libraryID;
    note.parentID = item.id;
    note.setNote(noteHtml);
    note.addTag(TABLE_NOTE_TAG);
    await note.saveTx();

    return note;
  }

  /**
   * 汇总多篇文献的表格内容，附加元数据供 LLM 引用
   *
   * 优化策略：表头只出现一次（置顶），后续每篇仅发送数据行，
   * 大幅减少 100+ 篇文献场景下的 token 消耗。
   *
   * @param tableResults 文献ID → 表格内容的映射
   * @param itemPdfPairs 父条目 → PDF 附件的映射（用于提取作者/年份）
   * @returns 合并后的 Markdown 文档
   */
  static aggregateTableContents(
    tableResults: Map<number, string>,
    itemPdfPairs?: Array<{
      parentItem: Zotero.Item;
      pdfAttachment: Zotero.Item;
    }>,
  ): string {
    // 构建 itemId → parentItem 的快速查找
    const itemMap = new Map<number, Zotero.Item>();
    if (itemPdfPairs) {
      for (const { parentItem } of itemPdfPairs) {
        itemMap.set(parentItem.id, parentItem);
      }
    }

    // 辅助函数：从 Markdown 表格中分离表头和数据行
    const splitTableHeaderAndRows = (
      md: string,
    ): { header: string; dataRows: string; nonTableContent: string } => {
      const lines = md.split("\n");
      const headerLines: string[] = [];
      const dataLines: string[] = [];
      const nonTableLines: string[] = [];
      let headerDone = false;

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed) continue;

        if (trimmed.startsWith("|")) {
          if (!headerDone) {
            headerLines.push(trimmed);
            // 分隔行（如 |---|---|---| ）标志表头结束
            if (/^\|[\s\-:|]+\|$/.test(trimmed)) {
              headerDone = true;
            }
          } else {
            dataLines.push(trimmed);
          }
        } else {
          nonTableLines.push(trimmed);
        }
      }

      return {
        header: headerLines.join("\n"),
        dataRows: dataLines.join("\n"),
        nonTableContent: nonTableLines.join("\n"),
      };
    };

    // 辅助函数：提取作者姓氏
    const extractAuthorSurname = (item: Zotero.Item): string => {
      const creators = (item as any).getCreators?.() || [];
      if (creators.length === 0) return "未知";
      const c = creators[0];
      if (c.lastName) return c.lastName;
      if (c.name) {
        const nameParts = c.name.trim().split(/\s+/);
        return nameParts[nameParts.length - 1];
      }
      return "未知";
    };

    // 辅助函数：提取年份
    const extractYear = (item: Zotero.Item): string => {
      const dateStr = (item.getField("date") as string) || "";
      const m = dateStr.match(/(\d{4})/);
      return m ? m[1] : "未知";
    };

    let globalHeader = "";
    const parts: string[] = [];
    let index = 1;

    for (const [itemId, tableContent] of tableResults) {
      const parentItem = itemMap.get(itemId);

      // 提取作者与年份标注
      let label: string;
      if (parentItem) {
        const author = extractAuthorSurname(parentItem);
        const year = extractYear(parentItem);
        const title = ((parentItem.getField("title") as string) || "").slice(
          0,
          80,
        );
        label = `> **文献 ${index}**: ${title} (${author}, ${year})`;
      } else {
        label = `> **文献 ${index}**`;
      }

      const { header, dataRows, nonTableContent } =
        splitTableHeaderAndRows(tableContent);

      if (!globalHeader && header) {
        // 首次遇到表头，记录为全局表头
        globalHeader = header;
      }

      // 组装：标注 + 数据行（无表头）
      let entry = label;
      if (nonTableContent) {
        entry += `\n${nonTableContent}`;
      }
      if (dataRows) {
        entry += `\n${dataRows}`;
      } else {
        // 如果没有解析出数据行（表格格式不标准），原样输出
        entry += `\n${tableContent}`;
      }

      parts.push(entry);
      index++;
    }

    // 拼装：全局表头 + 所有文献数据
    let result = "";
    if (globalHeader) {
      result += `**表格结构定义（以下每篇文献的数据行均遵循此表头）：**\n\n${globalHeader}\n\n---\n\n`;
    }
    result += parts.join("\n\n---\n\n");

    return result;
  }

  /**
   * 并行填表（带并发控制）
   *
   * @param items 文献条目与 PDF 附件的配对列表
   * @param tableTemplate 表格模板
   * @param fillPrompt 填表提示词
   * @param concurrency 并发数
   * @param progressCallback 进度回调 (done, total)
   * @returns 文献ID → 表格内容的映射
   */
  static async fillTablesInParallel(
    items: Array<{ parentItem: Zotero.Item; pdfAttachment: Zotero.Item }>,
    tableTemplate: string,
    fillPrompt: string,
    concurrency: number,
    progressCallback?: (done: number, total: number) => void,
  ): Promise<Map<number, string>> {
    const results = new Map<number, string>();
    let completed = 0;
    const total = items.length;
    const queue = [...items];

    const worker = async () => {
      while (queue.length > 0) {
        const task = queue.shift()!;
        try {
          // 先查缓存
          const existing = await this.findTableNote(task.parentItem);
          if (existing) {
            results.set(task.parentItem.id, existing);
          } else {
            const table = await this.fillTableForSinglePDF(
              task.parentItem,
              task.pdfAttachment,
              tableTemplate,
              fillPrompt,
            );
            await this.saveTableNote(task.parentItem, table);
            results.set(task.parentItem.id, table);
          }
        } catch (error) {
          ztoolkit.log(
            `[AI-Butler] 填表失败: ${task.parentItem.getField("title")}`,
            error,
          );
          results.set(
            task.parentItem.id,
            `(填表失败: ${error instanceof Error ? error.message : String(error)})`,
          );
        }
        completed++;
        progressCallback?.(completed, total);
      }
    };

    // 启动 N 个并行 worker
    const effectiveConcurrency = Math.min(concurrency, total);
    await Promise.all(
      Array.from({ length: effectiveConcurrency }, () => worker()),
    );

    return results;
  }

  /**
   * 创建报告条目
   */
  static async createReportItem(
    collection: Zotero.Collection,
    reportName: string,
  ): Promise<Zotero.Item> {
    const item = new Zotero.Item("report");
    item.setField("title", reportName);
    item.libraryID = collection.libraryID;

    // 使用事务包装保存和添加到分类操作
    await Zotero.DB.executeTransaction(async () => {
      await item.save();
      await collection.addItem(item.id);
    });

    return item;
  }

  /**
   * 将 PDF 附件添加到报告条目
   *
   * 创建链接附件，将原始 PDF 链接到报告条目下
   * 附件命名格式：论文标题前N位 + 原附件名称
   * 优化：缓存父条目标题，避免重复查询
   */
  static async attachPdfsToReport(
    reportItem: Zotero.Item,
    pdfAttachments: Zotero.Item[],
  ): Promise<void> {
    const TITLE_PREFIX_LENGTH = 30; // 论文标题前缀长度

    // 缓存父条目标题
    const parentTitleCache = new Map<number, string>();

    for (const pdfAtt of pdfAttachments) {
      try {
        // 获取原始 PDF 文件路径
        const filePath = await pdfAtt.getFilePathAsync();
        if (!filePath) {
          ztoolkit.log(`[AI-Butler] PDF 附件无文件路径: ${pdfAtt.id}`);
          continue;
        }

        // 获取原始附件的标题
        const originalTitle = (pdfAtt.getField("title") as string) || "PDF";

        // 获取父条目（论文）的标题（带缓存）
        let paperTitle = "";
        const parentID = pdfAtt.parentID;
        if (parentID) {
          if (parentTitleCache.has(parentID)) {
            paperTitle = parentTitleCache.get(parentID) || "";
          } else {
            const parentItem = await Zotero.Items.getAsync(parentID);
            if (parentItem) {
              paperTitle = (
                (parentItem.getField("title") as string) || ""
              ).trim();
              parentTitleCache.set(parentID, paperTitle);
            }
          }
        }

        // 构建新的附件标题：论文标题前N位 + 原附件名称
        let newTitle = originalTitle;
        if (paperTitle) {
          const titlePrefix =
            paperTitle.length > TITLE_PREFIX_LENGTH
              ? paperTitle.substring(0, TITLE_PREFIX_LENGTH) + "..."
              : paperTitle;
          newTitle = `[${titlePrefix}] ${originalTitle}`;
        }

        // 创建链接附件
        await Zotero.Attachments.linkFromFile({
          file: filePath,
          parentItemID: reportItem.id,
          title: newTitle,
        });
      } catch (error) {
        ztoolkit.log(`[AI-Butler] 添加 PDF 附件失败:`, error);
        // 继续处理其他附件
      }
    }
  }

  /**
   * 从 PDF 附件提取内容（包括文件路径）
   * 优化：缓存父条目信息，避免重复查询
   */
  static async extractPDFContentsFromAttachments(
    pdfAttachments: Zotero.Item[],
    progressCallback?: (message: string, progress: number) => void,
  ): Promise<PdfFileData[]> {
    const contents: PdfFileData[] = [];
    const total = pdfAttachments.length;

    // 缓存父条目标题，避免重复查询
    const parentTitleCache = new Map<number, string>();
    // 统计每个父条目有多少个 PDF，用于判断是否需要显示附件名
    const parentPdfCount = new Map<number, number>();

    // 第一遍：统计每个父条目的 PDF 数量
    for (const pdfAtt of pdfAttachments) {
      const parentID = pdfAtt.parentID;
      if (parentID) {
        parentPdfCount.set(parentID, (parentPdfCount.get(parentID) || 0) + 1);
      }
    }

    for (let i = 0; i < pdfAttachments.length; i++) {
      const pdfAtt = pdfAttachments[i];
      const attachmentTitle =
        (pdfAtt.getField("title") as string) || `PDF ${i + 1}`;
      const progress = 30 + Math.floor((i / total) * 20);
      progressCallback?.(
        `正在提取 (${i + 1}/${total}): ${attachmentTitle.slice(0, 30)}...`,
        progress,
      );

      try {
        // 获取文件路径
        const filePath = await pdfAtt.getFilePathAsync();
        if (!filePath) {
          ztoolkit.log(`[AI-Butler] PDF 附件无文件路径: ${pdfAtt.id}`);
          continue;
        }

        // 获取父条目标题（带缓存）
        let paperTitle = "";
        const parentID = pdfAtt.parentID;
        if (parentID) {
          if (parentTitleCache.has(parentID)) {
            paperTitle = parentTitleCache.get(parentID) || "";
          } else {
            const parentItem = await Zotero.Items.getAsync(parentID);
            if (parentItem) {
              paperTitle = (
                (parentItem.getField("title") as string) || ""
              ).trim();
              parentTitleCache.set(parentID, paperTitle);
            }
          }
        }

        // 构建显示标题：如果同一论文有多个 PDF，则显示 "论文标题 - 附件名"
        let displayTitle = paperTitle || attachmentTitle;
        const pdfCountForParent = parentID
          ? parentPdfCount.get(parentID) || 1
          : 1;
        if (pdfCountForParent > 1 && paperTitle) {
          displayTitle = `${paperTitle} - ${attachmentTitle}`;
        }

        // 尝试读取 Base64 内容
        let base64Content = "";
        try {
          const fileData = await IOUtils.read(filePath);
          // 使用分块方式转换为 base64，避免大文件导致 "too many function arguments" 错误
          base64Content = this.arrayBufferToBase64(fileData);
        } catch (e) {
          ztoolkit.log(`[AI-Butler] 读取 PDF 文件失败: ${filePath}`, e);
        }

        contents.push({
          title: displayTitle,
          filePath,
          content: base64Content,
          isBase64: true,
        });
      } catch (error) {
        ztoolkit.log(
          `[AI-Butler] 提取 PDF 内容失败: ${attachmentTitle}`,
          error,
        );
        // 继续处理其他文献
      }
    }

    return contents;
  }

  /**
   * 使用 LLM 从多个 PDF 生成综述
   */
  static async generateSummaryFromMultiplePDFs(
    pdfContents: PdfFileData[],
    prompt: string,
    progressCallback?: (message: string, progress: number) => void,
  ): Promise<string> {
    if (pdfContents.length === 0) {
      throw new Error("没有可用的 PDF 内容");
    }

    // 检查当前使用的 API 提供商
    const providerName = (getPref("provider") as string) || "google";
    const provider = ProviderRegistry.get(providerName);

    // 检查 provider 是否支持多文件处理
    const supportsMultiFile =
      provider && typeof provider.generateMultiFileSummary === "function";

    // 判断是否是 Gemini 提供商（支持 google 和 gemini 两种名称）
    const isGemini =
      providerName === "google" ||
      providerName.toLowerCase().includes("gemini");

    if (supportsMultiFile && isGemini) {
      // 使用 Gemini 多文件模式 (inline_data)
      return await this.generateWithGeminiFileAPI(
        pdfContents,
        prompt,
        progressCallback,
      );
    } else {
      // 回退到合并文本模式
      return await this.generateWithMergedText(
        pdfContents,
        prompt,
        progressCallback,
      );
    }
  }

  /**
   * 使用 Gemini File API 生成综述
   */
  private static async generateWithGeminiFileAPI(
    pdfContents: PdfFileData[],
    prompt: string,
    progressCallback?: (message: string, progress: number) => void,
  ): Promise<string> {
    progressCallback?.("正在上传 PDF 文件到 Gemini...", 55);

    // 获取 Gemini provider（支持 google 和 gemini 两种名称）
    let provider = ProviderRegistry.get("google");
    if (!provider) {
      provider = ProviderRegistry.get("gemini");
    }
    if (!provider || typeof provider.generateMultiFileSummary !== "function") {
      throw new Error("Gemini provider 不支持多文件处理");
    }

    // 构建 PDF 文件信息列表
    const pdfFiles: PdfFileInfo[] = pdfContents.map((pdf, index) => ({
      filePath: pdf.filePath,
      displayName: `${index + 1}_${pdf.title.slice(0, 50)}`,
      base64Content: pdf.content,
    }));

    // 获取 LLM 选项
    const options = LLMClient.getLLMOptions();

    progressCallback?.("正在调用 AI 生成综述...", 65);

    // 调用 Gemini 多文件处理
    const result = await provider.generateMultiFileSummary(
      pdfFiles,
      prompt,
      options,
    );

    return result;
  }

  /**
   * 使用合并文本模式生成综述
   */
  private static async generateWithMergedText(
    pdfContents: PdfFileData[],
    prompt: string,
    progressCallback?: (message: string, progress: number) => void,
  ): Promise<string> {
    progressCallback?.("正在调用 AI 生成综述 (文本模式)...", 60);

    // 如果有 Base64 内容但 provider 不支持多文件，尝试提取文本
    let combinedContent = "";
    let hasBase64 = false;
    let firstBase64Content = "";

    for (const pdf of pdfContents) {
      if (pdf.isBase64 && pdf.content) {
        if (!hasBase64) {
          hasBase64 = true;
          firstBase64Content = pdf.content;
        }
        combinedContent += `\n\n=== 论文: ${pdf.title} ===\n[PDF 内容]\n`;
      } else {
        combinedContent += `\n\n=== 论文: ${pdf.title} ===\n${pdf.content}\n`;
      }
    }

    // 如果有 Base64 内容，使用第一个 PDF 的 Base64
    if (hasBase64 && firstBase64Content) {
      const fullPrompt = `${prompt}\n\n以下是需要综述的论文列表:\n${pdfContents.map((p, i) => `${i + 1}. ${p.title}`).join("\n")}\n\n请基于上传的 PDF 内容生成综述。`;

      const result = await LLMClient.generateSummaryWithRetry(
        firstBase64Content,
        true,
        fullPrompt,
      );
      return result;
    }

    // 纯文本模式
    if (!combinedContent.trim()) {
      throw new Error("当前 API 不支持多文件处理，且无法提取 PDF 文本内容");
    }

    const fullPrompt = `${prompt}\n\n以下是需要综述的论文内容:\n${combinedContent}`;

    const result = await LLMClient.generateSummaryWithRetry(
      combinedContent,
      false,
      fullPrompt,
    );

    return result;
  }

  /**
   * 创建综述笔记（兼容旧接口，用于子笔记创建）
   */
  static async createReviewNote(
    reportItem: Zotero.Item,
    reviewName: string,
    content: string,
  ): Promise<Zotero.Item> {
    const formattedContent = NoteGenerator.formatNoteContent(
      reviewName,
      content,
    );
    const note = await NoteGenerator.createNote(reportItem, formattedContent);
    return note;
  }

  /**
   * 创建独立综述笔记（直接放在分类目录下，无父条目）
   *
   * @param collection 目标分类
   * @param reviewName 综述名称
   * @param content 综述正文（Markdown）
   * @returns 创建的笔记条目
   */
  static async createStandaloneReviewNote(
    collection: Zotero.Collection,
    reviewName: string,
    content: string,
  ): Promise<Zotero.Item> {
    // 格式化内容
    const formattedContent = NoteGenerator.formatNoteContent(
      reviewName,
      content,
    );

    // 创建独立笔记（无父条目）
    const note = new Zotero.Item("note");
    note.libraryID = collection.libraryID;
    note.setNote(formattedContent);
    note.addTag("AI-Review");

    // 保存并添加到分类
    await Zotero.DB.executeTransaction(async () => {
      await note.save();
      await collection.addItem(note.id);
    });

    return note;
  }

  /**
   * 后处理综述正文中的引用标记，转换为 Zotero 链接
   *
   * 匹配 LLM 自然生成的 (Author, Year) 格式引用，
   * 基于文献元数据（作者姓氏/年份）进行模糊匹配，
   * 将匹配成功的引用转换为 zotero://select 可点击链接。
   *
   * @param content 综述正文
   * @param itemPdfPairs 文献条目列表
   * @returns 处理后的正文
   */
  static async postProcessCitations(
    content: string,
    itemPdfPairs: Array<{
      parentItem: Zotero.Item;
      pdfAttachment: Zotero.Item;
    }>,
  ): Promise<string> {
    // 构建作者+年份 → item 的查找表
    // key 格式: "surname|year" (小写)
    const authorYearMap = new Map<
      string,
      { item: Zotero.Item; key: string; uri: string }
    >();

    for (const { parentItem } of itemPdfPairs) {
      const creators = (parentItem as any).getCreators?.() || [];
      const itemKey = (parentItem as any).key || "";
      const uri = `zotero://select/library/items/${itemKey}`;
      const dateStr = (parentItem.getField("date") as string) || "";
      const yearMatch = dateStr.match(/(\d{4})/);
      const year = yearMatch ? yearMatch[1] : "";

      if (!year || creators.length === 0) continue;

      // 注册所有作者的姓氏（支持多作者匹配）
      for (const creator of creators) {
        let surname = "";
        if (creator.lastName) {
          surname = creator.lastName.trim();
        } else if (creator.name) {
          // 单字段格式如 "F. Begarin"，取最后一个词作为姓氏
          const nameParts = creator.name.trim().split(/\s+/);
          surname = nameParts[nameParts.length - 1];
        }
        if (!surname) continue;

        const lookupKey = `${surname.toLowerCase()}|${year}`;
        if (!authorYearMap.has(lookupKey)) {
          authorYearMap.set(lookupKey, { item: parentItem, key: itemKey, uri });
        }
      }
    }

    if (authorYearMap.size === 0) return content;

    // 匹配 (Author, Year)、(Author et al., Year)、(Author and Author, Year) 等
    // 正则: 括号内以字母开头，包含逗号分隔的年份
    let result = content;
    result = result.replace(
      /\(([^()]{2,80}?,\s*\d{4}[a-z]?)\)/g,
      (fullMatch, inner: string) => {
        // 提取年份
        const yearMatch = inner.match(/(\d{4})[a-z]?\s*$/);
        if (!yearMatch) return fullMatch;
        const year = yearMatch[1];

        // 提取作者部分（逗号前面的内容）
        const authorPart = inner.replace(/,\s*\d{4}[a-z]?\s*$/, "").trim();

        // 尝试从作者部分提取姓氏
        // 处理 "Author et al." → "Author"
        // 处理 "Author and Author" → 取第一个
        // 处理 "Author" → 直接使用
        let surname = authorPart
          .replace(/\s+et\s+al\.?$/i, "")
          .replace(/\s+and\s+.+$/i, "")
          .replace(/\s+&\s+.+$/i, "")
          .trim();

        // 如果包含空格，取最后一个词作为姓氏（"First Last" → "Last"）
        // 但如果是单个词则直接使用
        const parts = surname.split(/\s+/);
        if (parts.length > 1) {
          surname = parts[parts.length - 1];
        }

        const lookupKey = `${surname.toLowerCase()}|${year}`;
        const match = authorYearMap.get(lookupKey);

        if (match) {
          return `[(${inner})](${match.uri})`;
        }

        return fullMatch;
      },
    );

    // 模式2: Author (Year) / Author et al. (Year) / Author and Author (Year)
    // 叙述性引用格式，如 "Nicoleau (2014)" "Nicoleau et al. (2014)"
    result = result.replace(
      /(?<!\[)\b([A-Z][a-zA-Zà-öø-ÿÀ-ÖØ-Ý\-']+(?:\s+(?:et\s+al\.?|and|&)\s+[A-Za-zà-öø-ÿÀ-ÖØ-Ý\-']+)?)\s+\((\d{4}[a-z]?)\)(?!\])/g,
      (fullMatch, authorText: string, yearWithSuffix: string) => {
        const year = yearWithSuffix.slice(0, 4);

        // 提取第一作者姓氏
        let surname = authorText
          .replace(/\s+et\s+al\.?$/i, "")
          .replace(/\s+(and|&)\s+.+$/i, "")
          .trim();

        // 取最后一个词作为姓氏
        const parts = surname.split(/\s+/);
        if (parts.length > 1) {
          surname = parts[parts.length - 1];
        }

        const lookupKey = `${surname.toLowerCase()}|${year}`;
        const match = authorYearMap.get(lookupKey);

        if (match) {
          return `[${authorText} (${yearWithSuffix})](${match.uri})`;
        }

        return fullMatch;
      },
    );

    // 清理残留的 [itemId:N] 标记（如果 LLM 仍然生成了的话）
    result = result.replace(/\[itemId:\d+\]/g, "");

    return result;
  }

  /**
   * 将 ArrayBuffer 转换为 Base64 字符串
   * 使用分块处理避免 "too many function arguments" 错误
   */
  private static arrayBufferToBase64(buffer: ArrayBuffer | Uint8Array): string {
    const bytes =
      buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const chunkSize = 0x8000; // 32KB chunks
    let result = "";

    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));
      result += String.fromCharCode.apply(null, Array.from(chunk));
    }

    return btoa(result);
  }
}
