const ITEMS_PER_PAGE = 2;
const PDF_PAGE_WIDTH_PT = 595.28;
const PDF_PAGE_HEIGHT_PT = 841.89;
const EXPORT_RENDER_SCALE = Math.max(2, window.devicePixelRatio || 1);
const A4_RATIO = 297 / 210;
const A4_HEIGHT_TOLERANCE_PX = 6;
const DETAIL_TEXT_CHUNK_SIZE = 118;
const DETAIL_FIRST_PAGE_CAPACITY = 170;
const DETAIL_CONTINUED_PAGE_CAPACITY = 320;

const documentRootElement = document.getElementById("document-root");
const downloadAllButtonElement = document.getElementById("download-all-pdf");
const downloadPageButtonElement = document.getElementById("download-page-pdf");
const downloadPageSelectElement = document.getElementById("download-page-select");
const exportStatusElement = document.getElementById("export-status");
const diagnosticPanelElement = document.getElementById("diagnostic-panel");
const diagnosticSummaryElement = document.getElementById("diagnostic-summary");
const diagnosticCodeElement = document.getElementById("diagnostic-code");
const copyLayoutCodeButtonElement = document.getElementById("copy-layout-code");

let exportInProgress = false;
let currentPortfolio = null;
let currentLayoutDiagnostics = [];
let resizeDiagnosticTimer = 0;

async function loadJson(path) {
  try {
    const response = await fetch(path, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    return await response.json();
  } catch (error) {
    if (window.location.protocol !== "file:") {
      throw error;
    }
    return loadJsonFromIframe(path);
  }
}

function loadJsonFromIframe(path) {
  return new Promise((resolve, reject) => {
    const frame = document.createElement("iframe");
    frame.hidden = true;
    frame.src = path;

    const cleanup = () => {
      frame.onload = null;
      frame.onerror = null;
      frame.remove();
    };

    frame.onload = () => {
      try {
        const text = frame.contentDocument?.body?.textContent?.trim();
        cleanup();

        if (!text) {
          throw new Error("Empty response");
        }

        resolve(JSON.parse(text));
      } catch (parseError) {
        reject(parseError);
      }
    };

    frame.onerror = () => {
      cleanup();
      reject(new Error(`Unable to load ${path}`));
    };

    document.body.appendChild(frame);
  });
}

function chunkItems(items, size) {
  const chunks = [];

  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size));
  }

  return chunks;
}

function getSectionMarker(title) {
  const match = title.match(/^([一二三四五六七八九十]+、)/);
  return match ? match[1] : "";
}

function getSectionTitleText(title) {
  return title.replace(/^[一二三四五六七八九十]+、\s*/, "").trim();
}

function hasDetailSections(item) {
  return Array.isArray(item.detailSections) && item.detailSections.length > 0;
}

function splitParagraphIntoChunks(paragraph, chunkSize = DETAIL_TEXT_CHUNK_SIZE) {
  const normalized = String(paragraph || "").trim();

  if (!normalized) {
    return [];
  }

  const primaryParts = normalized.match(/[^。！？；]+[。！？；]?/g) || [normalized];
  const chunks = [];
  let current = "";

  primaryParts
    .map((part) => part.trim())
    .filter(Boolean)
    .forEach((part) => {
      if (!current || current.length + part.length <= chunkSize) {
        current += part;
        return;
      }

      chunks.push(current);
      current = part;
    });

  if (current) {
    chunks.push(current);
  }

  return chunks.flatMap((chunk) => {
    if (chunk.length <= chunkSize * 1.2) {
      return [chunk];
    }

    const refinedParts = chunk.match(/[^，、]+[，、]?/g) || [chunk];
    const refinedChunks = [];
    let refinedCurrent = "";

    refinedParts
      .map((part) => part.trim())
      .filter(Boolean)
      .forEach((part) => {
        if (!refinedCurrent || refinedCurrent.length + part.length <= chunkSize) {
          refinedCurrent += part;
          return;
        }

        refinedChunks.push(refinedCurrent);
        refinedCurrent = part;
      });

    if (refinedCurrent) {
      refinedChunks.push(refinedCurrent);
    }

    return refinedChunks;
  });
}

function createDetailBlocks(item) {
  return (item.detailSections || []).flatMap((detailSection) => {
    const paragraphs = (detailSection.paragraphs || []).flatMap((paragraph) =>
      splitParagraphIntoChunks(paragraph)
    );

    return paragraphs.map((paragraph, index) => ({
      heading: index === 0 ? detailSection.heading : `${detailSection.heading}（續）`,
      paragraphs: [paragraph]
    }));
  });
}

function getDetailBlockWeight(block) {
  const paragraphLength = (block.paragraphs || []).join("").length;
  return paragraphLength + Math.max(28, (block.heading || "").length * 4);
}

function distributeDetailBlocks(blocks, pageCount, imageCount, descLength) {
  const resolvedPageCount = Math.max(1, Math.min(pageCount, blocks.length || 1));
  const capacities = Array.from({ length: resolvedPageCount }, (_, index) => {
    if (index === 0) {
      return Math.max(
        150,
        DETAIL_FIRST_PAGE_CAPACITY - imageCount * 32 - Math.round(descLength * 0.22)
      );
    }

    return DETAIL_CONTINUED_PAGE_CAPACITY;
  });

  const pages = Array.from({ length: resolvedPageCount }, () => []);
  let pageIndex = 0;
  let pageWeight = 0;

  blocks.forEach((block, blockIndex) => {
    const weight = getDetailBlockWeight(block);
    const remainingBlocks = blocks.length - blockIndex;
    const remainingPages = resolvedPageCount - pageIndex;
    const shouldAdvance =
      pageIndex < resolvedPageCount - 1 &&
      pages[pageIndex].length > 0 &&
      (pageWeight + weight > capacities[pageIndex] || remainingBlocks === remainingPages);

    if (shouldAdvance) {
      pageIndex += 1;
      pageWeight = 0;
    }

    pages[pageIndex].push(block);
    pageWeight += weight;
  });

  return pages.filter((page) => page.length > 0);
}

function buildDetailedItemParts(item, itemNumber, imageLibrary, requestedParts) {
  const images = imageLibrary[item.id] || [];
  const detailBlocks = createDetailBlocks(item);
  const parts = distributeDetailBlocks(detailBlocks, requestedParts, images.length, item.desc.length);
  const useDedicatedEvidencePage = parts.length > 1 && images.length === 1;
  const partTotal = parts.length + (useDedicatedEvidencePage ? 1 : 0);
  const baseParts = parts.map((detailSections, index) => ({
    kind: "detail",
    item,
    itemNumber,
    detailPage: {
      pageIndex: index + 1,
      pageTotal: partTotal,
      detailSections,
      desc: index === 0 ? item.desc : "",
      descLabel: index === 0 ? "內容摘要" : "續頁內容",
      images: useDedicatedEvidencePage ? [] : index === 0 ? images : [],
      showEvidence: !useDedicatedEvidencePage && index === 0 && images.length > 0,
      evidenceOnly: false
    }
  }));

  if (!useDedicatedEvidencePage) {
    return baseParts;
  }

  return [
    ...baseParts,
    {
      kind: "detail",
      item,
      itemNumber,
      detailPage: {
        pageIndex: partTotal,
        pageTotal: partTotal,
        detailSections: [],
        desc: "",
        descLabel: "佐證資料",
        images,
        showEvidence: true,
        evidenceOnly: true
      }
    }
  ];
}

function buildSectionParts(section, imageLibrary, splitOverrides = {}) {
  const parts = [];
  let briefItems = [];

  section.items.forEach((item, itemIndex) => {
    const itemNumber = itemIndex + 1;

    if (hasDetailSections(item)) {
      if (briefItems.length) {
        parts.push(briefItems);
        briefItems = [];
      }

      parts.push(
        ...buildDetailedItemParts(item, itemNumber, imageLibrary, splitOverrides[item.id] || 1)
      );
      return;
    }

    briefItems.push({
      kind: "brief",
      item,
      itemNumber
    });

    if (briefItems.length === ITEMS_PER_PAGE) {
      parts.push(briefItems);
      briefItems = [];
    }
  });

  if (briefItems.length) {
    parts.push(briefItems);
  }

  return parts;
}

function sanitizeFileName(value) {
  return String(value || "")
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

function getDocumentPages() {
  return Array.from(documentRootElement.querySelectorAll(".a4-page"));
}

function getDocumentBaseName() {
  const parts = [currentPortfolio?.title, currentPortfolio?.theme].filter(Boolean);
  return sanitizeFileName(parts.join("_")) || "多元表現";
}

function getPageExportLabel(pageElement, pageIndex) {
  const explicitLabel = pageElement.dataset.exportLabel?.trim();

  if (explicitLabel) {
    return explicitLabel;
  }

  const code = pageElement.querySelector(".page-code")?.textContent?.trim();
  return code || `第 ${pageIndex + 1} 頁`;
}

function getPageFileName(pageElement, pageIndex) {
  const pageNumber = String(pageIndex + 1).padStart(2, "0");
  const label = sanitizeFileName(getPageExportLabel(pageElement, pageIndex));
  return `${getDocumentBaseName()}_第${pageNumber}頁_${label}.pdf`;
}

function getDocumentFileName() {
  return `${getDocumentBaseName()}_完整.pdf`;
}

function setExportStatus(message = "", tone = "") {
  exportStatusElement.textContent = message;

  if (tone) {
    exportStatusElement.dataset.tone = tone;
    return;
  }

  delete exportStatusElement.dataset.tone;
}

function refreshExportControls() {
  const hasPages = Boolean(currentPortfolio) && getDocumentPages().length > 0;

  downloadAllButtonElement.disabled = exportInProgress || !hasPages;
  downloadPageSelectElement.disabled = exportInProgress || !hasPages;
  downloadPageButtonElement.disabled =
    exportInProgress ||
    !hasPages ||
    !downloadPageSelectElement.value;
}

function getExpectedPageHeightPx(pageElement) {
  const rect = pageElement.getBoundingClientRect();
  return rect.width * A4_RATIO;
}

function getPageItemLabels(pageElement) {
  if (pageElement.classList.contains("cover-page")) {
    return ["封面"];
  }

  const labels = Array.from(pageElement.querySelectorAll(".item-name"))
    .map((node) => node.textContent?.trim())
    .filter(Boolean);

  if (labels.length) {
    return labels;
  }

  const pageNumber = Number(pageElement.dataset.pageNumber || "1");
  return [getPageExportLabel(pageElement, Math.max(0, pageNumber - 1))];
}

function buildLayoutDiagnostics(pageElements = getDocumentPages()) {
  return pageElements
    .map((pageElement, pageIndex) => {
      const pageNumber = Number(pageElement.dataset.pageNumber || pageIndex + 1);
      const actualHeightPx = Math.round(pageElement.getBoundingClientRect().height);
      const expectedHeightPx = Math.round(getExpectedPageHeightPx(pageElement));
      const overflowPx = actualHeightPx - expectedHeightPx;

      if (overflowPx <= A4_HEIGHT_TOLERANCE_PX) {
        return null;
      }

      const pageLabel = getPageExportLabel(pageElement, pageNumber - 1);
      const itemLabels = getPageItemLabels(pageElement);

      return {
        pageNumber,
        pageLabel,
        itemLabels,
        expectedHeightPx,
        actualHeightPx,
        overflowPx,
        overflowRatio: Number((actualHeightPx / expectedHeightPx).toFixed(4)),
        code: `E-A4-HEIGHT-OVERFLOW-P${String(pageNumber).padStart(2, "0")}`,
        type: "a4_height_overflow",
        cause: "page_content_exceeds_a4_height",
        effect: "pdf_page_will_be_vertically_compressed"
      };
    })
    .filter(Boolean);
}

function buildLayoutDiagnosticPayload(diagnostics) {
  return JSON.stringify(
    {
      version: "PORTFOLIO_LAYOUT_WARN_V1",
      document: getDocumentBaseName(),
      issueCount: diagnostics.length,
      issues: diagnostics.map((issue) => ({
        page: issue.pageNumber,
        code: issue.code,
        type: issue.type,
        label: issue.pageLabel,
        items: issue.itemLabels,
        expectedHeightPx: issue.expectedHeightPx,
        actualHeightPx: issue.actualHeightPx,
        overflowPx: issue.overflowPx,
        overflowRatio: issue.overflowRatio,
        cause: issue.cause,
        effect: issue.effect
      }))
    },
    null,
    2
  );
}

function clearPageLayoutWarnings() {
  getDocumentPages().forEach((pageElement) => {
    pageElement.removeAttribute("data-layout-warning");
    pageElement.removeAttribute("data-warning-code");
    pageElement.querySelector(".page-layout-warning")?.remove();
  });
}

function renderPageLayoutWarnings(diagnostics) {
  clearPageLayoutWarnings();

  const pageMap = new Map(getDocumentPages().map((pageElement) => [Number(pageElement.dataset.pageNumber), pageElement]));

  diagnostics.forEach((diagnostic) => {
    const pageElement = pageMap.get(diagnostic.pageNumber);

    if (!pageElement) {
      return;
    }

    pageElement.dataset.layoutWarning = "true";
    pageElement.dataset.warningCode = diagnostic.code;

    const warning = document.createElement("aside");
    warning.className = "page-layout-warning pdf-control";
    warning.setAttribute("role", "note");
    warning.title = `${diagnostic.code}：此頁實際高度超出 A4 ${diagnostic.overflowPx}px，若直接匯出 PDF 會被壓扁。`;
    warning.innerHTML = `
      <span class="page-layout-warning__label">PDF 匯出警告</span>
      <strong class="page-layout-warning__code">${diagnostic.code}</strong>
      <span class="page-layout-warning__meta">高度超出 A4 ${diagnostic.overflowPx}px</span>
    `;

    pageElement.appendChild(warning);
  });
}

function renderLayoutDiagnostics(diagnostics) {
  currentLayoutDiagnostics = diagnostics;
  renderPageLayoutWarnings(diagnostics);

  if (!diagnosticPanelElement || !diagnosticSummaryElement || !diagnosticCodeElement || !copyLayoutCodeButtonElement) {
    return;
  }

  if (!diagnostics.length) {
    diagnosticPanelElement.hidden = true;
    diagnosticSummaryElement.textContent = "";
    diagnosticCodeElement.textContent = "";
    copyLayoutCodeButtonElement.disabled = true;
    return;
  }

  diagnosticPanelElement.hidden = false;
  copyLayoutCodeButtonElement.disabled = false;

  const pageList = diagnostics.map((issue) => issue.pageNumber).join("、");
  diagnosticSummaryElement.textContent =
    `下列頁面實際高度已超出 A4：${pageList}。若直接匯出，整頁會被垂直壓縮。請先縮短內容或調整排版；也可將下方錯誤代碼交給 Codex 修正。`;
  diagnosticCodeElement.textContent = buildLayoutDiagnosticPayload(diagnostics);
}

async function waitForLayoutStability() {
  if (document.fonts?.ready) {
    await document.fonts.ready;
  }

  await new Promise((resolve) => window.requestAnimationFrame(() => window.requestAnimationFrame(resolve)));
}

async function refreshLayoutDiagnostics() {
  if (!currentPortfolio) {
    renderLayoutDiagnostics([]);
    return [];
  }

  await waitForLayoutStability();
  const diagnostics = buildLayoutDiagnostics(getDocumentPages());
  renderLayoutDiagnostics(diagnostics);
  return diagnostics;
}

function getDiagnosticsForPages(pageElements, diagnostics = currentLayoutDiagnostics) {
  const pageNumbers = new Set(
    pageElements
      .map((pageElement) => Number(pageElement.dataset.pageNumber))
      .filter((pageNumber) => Number.isFinite(pageNumber))
  );

  return diagnostics.filter((issue) => pageNumbers.has(issue.pageNumber));
}

function focusDiagnosticPanel() {
  if (!diagnosticPanelElement?.hidden) {
    diagnosticPanelElement.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }

  const firstWarningPage = documentRootElement.querySelector('.a4-page[data-layout-warning="true"]');
  firstWarningPage?.scrollIntoView({ behavior: "smooth", block: "center" });
}

async function copyDiagnosticCode() {
  const text = diagnosticCodeElement?.textContent?.trim();

  if (!text) {
    setExportStatus("目前沒有可複製的版面錯誤代碼", "");
    return;
  }

  try {
    if (navigator.clipboard?.writeText && window.isSecureContext) {
      await navigator.clipboard.writeText(text);
    } else {
      const helper = document.createElement("textarea");
      helper.value = text;
      helper.setAttribute("readonly", "readonly");
      helper.style.position = "fixed";
      helper.style.opacity = "0";
      document.body.appendChild(helper);
      helper.select();
      document.execCommand("copy");
      helper.remove();
    }

    setExportStatus("已複製版面錯誤代碼", "success");
  } catch (error) {
    console.error(error);
    setExportStatus("錯誤代碼複製失敗，請手動複製下方內容", "error");
  }
}

function scheduleLayoutDiagnosticRefresh() {
  window.clearTimeout(resizeDiagnosticTimer);
  resizeDiagnosticTimer = window.setTimeout(() => {
    refreshLayoutDiagnostics().catch((error) => console.error(error));
  }, 180);
}

function updateExportOptions() {
  if (!currentPortfolio) {
    downloadPageSelectElement.innerHTML = "";

    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "請先載入內容";
    downloadPageSelectElement.appendChild(placeholder);
    refreshExportControls();
    return;
  }

  const pages = getDocumentPages();
  const previousValue = downloadPageSelectElement.value;
  downloadPageSelectElement.innerHTML = "";

  if (!pages.length) {
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "請先載入內容";
    downloadPageSelectElement.appendChild(placeholder);
    refreshExportControls();
    return;
  }

  const defaultOption = document.createElement("option");
  defaultOption.value = "";
  defaultOption.textContent = "選擇要下載的頁面";
  downloadPageSelectElement.appendChild(defaultOption);

  pages.forEach((pageElement, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = `第 ${index + 1} 頁｜${getPageExportLabel(pageElement, index)}`;
    downloadPageSelectElement.appendChild(option);
  });

  if (previousValue && pages[Number(previousValue)]) {
    downloadPageSelectElement.value = previousValue;
  }

  refreshExportControls();
}

function triggerFileDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();

  window.setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1500);
}

function waitForImageLoad(image) {
  if (image.complete && image.naturalWidth > 0) {
    return Promise.resolve();
  }

  return new Promise((resolve) => {
    const done = () => {
      image.removeEventListener("load", done);
      image.removeEventListener("error", done);
      resolve();
    };

    image.addEventListener("load", done, { once: true });
    image.addEventListener("error", done, { once: true });
  });
}

async function waitForPageAssets(pageElement) {
  if (document.fonts?.ready) {
    await document.fonts.ready;
  }

  const images = Array.from(pageElement.querySelectorAll("img"));
  await Promise.all(images.map((image) => waitForImageLoad(image)));
}

async function renderPageToCanvas(pageElement) {
  await waitForPageAssets(pageElement);

  if (typeof window.html2canvas !== "function") {
    throw new Error("PDF 匯出元件載入失敗，請重新整理後再試。");
  }

  const rect = pageElement.getBoundingClientRect();
  const width = Math.ceil(rect.width);
  const height = Math.ceil(rect.height);

  return window.html2canvas(pageElement, {
    backgroundColor: "#ffffff",
    useCORS: true,
    allowTaint: false,
    logging: false,
    scale: EXPORT_RENDER_SCALE,
    width,
    height,
    scrollX: 0,
    scrollY: -window.scrollY,
    onclone: (clonedDocument) => {
      clonedDocument.querySelectorAll(".pdf-control").forEach((node) => node.remove());

      const pageNumber = pageElement.dataset.pageNumber;
      const clonedPage = clonedDocument.querySelector(`.a4-page[data-page-number="${pageNumber}"]`);

      if (clonedPage) {
        clonedPage.style.width = `${width}px`;
        clonedPage.style.minWidth = `${width}px`;
        clonedPage.style.minHeight = `${height}px`;
        clonedPage.style.margin = "0";
        clonedPage.style.boxShadow = "none";
      }

      const sourceFrames = Array.from(pageElement.querySelectorAll(".cover-photo-frame, .image-frame"));
      const clonedFrames = Array.from(clonedDocument.querySelectorAll(`.a4-page[data-page-number="${pageNumber}"] .cover-photo-frame, .a4-page[data-page-number="${pageNumber}"] .image-frame`));

      sourceFrames.forEach((sourceFrame, index) => {
        const clonedFrame = clonedFrames[index];

        if (!clonedFrame) {
          return;
        }

        const frameRect = sourceFrame.getBoundingClientRect();
        const sourceImage = sourceFrame.querySelector("img");
        const clonedImage = clonedFrame.querySelector("img");

        clonedFrame.style.width = `${frameRect.width}px`;
        clonedFrame.style.height = `${frameRect.height}px`;
        clonedFrame.style.minWidth = `${frameRect.width}px`;
        clonedFrame.style.minHeight = `${frameRect.height}px`;
        clonedFrame.style.overflow = "hidden";

        if (!sourceImage || !clonedImage) {
          return;
        }

        const computedStyle = window.getComputedStyle(sourceImage);
        const imageUrl = sourceImage.currentSrc || sourceImage.src;

        clonedFrame.style.backgroundImage = `url("${imageUrl}")`;
        clonedFrame.style.backgroundRepeat = "no-repeat";
        clonedFrame.style.backgroundPosition = computedStyle.objectPosition || "center center";
        clonedFrame.style.backgroundSize = computedStyle.objectFit === "contain" ? "contain" : "cover";

        clonedImage.style.opacity = "0";
        clonedImage.style.width = `${frameRect.width}px`;
        clonedImage.style.height = `${frameRect.height}px`;
        clonedImage.style.minWidth = `${frameRect.width}px`;
        clonedImage.style.minHeight = `${frameRect.height}px`;
        clonedImage.style.objectFit = computedStyle.objectFit || "cover";
        clonedImage.style.objectPosition = computedStyle.objectPosition || "center center";
      });

      clonedDocument.body.style.background = "#ffffff";
    }
  });
}

function createPdfDocument() {
  const jsPdfConstructor = window.jspdf?.jsPDF;

  if (typeof jsPdfConstructor !== "function") {
    throw new Error("PDF 匯出元件載入失敗，請重新整理後再試。");
  }

  return new jsPdfConstructor({
    orientation: "portrait",
    unit: "pt",
    format: "a4",
    compress: true
  });
}

function appendCanvasToPdf(pdf, canvas, pageIndex) {
  if (pageIndex > 0) {
    pdf.addPage("a4", "portrait");
  }

  const imageData = canvas.toDataURL("image/jpeg", 0.96);
  pdf.addImage(imageData, "JPEG", 0, 0, PDF_PAGE_WIDTH_PT, PDF_PAGE_HEIGHT_PT, undefined, "FAST");
}

async function exportPagesToPdf(pageElements, filename, label) {
  const diagnostics = await refreshLayoutDiagnostics();
  const blockingIssues = getDiagnosticsForPages(pageElements, diagnostics);

  if (blockingIssues.length) {
    const pageList = blockingIssues.map((issue) => issue.pageNumber).join("、");
    setExportStatus(
      `${label}已停止：第 ${pageList} 頁高度超出 A4，直接匯出會被壓扁。請先修正頁面內容，或將錯誤代碼提供給 Codex。`,
      "error"
    );
    focusDiagnosticPanel();
    return;
  }

  exportInProgress = true;
  refreshExportControls();

  try {
    const pdf = createPdfDocument();

    for (let index = 0; index < pageElements.length; index += 1) {
      setExportStatus(`${label}：正在處理第 ${index + 1} / ${pageElements.length} 頁`, "");

      const canvas = await renderPageToCanvas(pageElements[index]);
      appendCanvasToPdf(pdf, canvas, index);
      canvas.width = 0;
      canvas.height = 0;
    }

    const pdfBlob = pdf.output("blob");
    triggerFileDownload(pdfBlob, filename);
    setExportStatus(`${label}：已開始下載`, "success");
  } catch (error) {
    console.error(error);
    setExportStatus(error.message || "PDF 匯出失敗", "error");
  } finally {
    exportInProgress = false;
    refreshExportControls();
  }
}

function handleExportAllPdf() {
  const pages = getDocumentPages();

  if (!pages.length) {
    return;
  }

  exportPagesToPdf(pages, getDocumentFileName(), "整份 PDF");
}

function handleExportSinglePagePdf() {
  const pageIndex = Number(downloadPageSelectElement.value);
  const pages = getDocumentPages();
  const targetPage = pages[pageIndex];

  if (!targetPage) {
    return;
  }

  exportPagesToPdf(
    [targetPage],
    getPageFileName(targetPage, pageIndex),
    `第 ${pageIndex + 1} 頁 PDF`
  );
}

function createPage(pageNumber, totalPages, footerText, pageClass = "") {
  const page = document.createElement("section");
  page.className = `a4-page ${pageClass}`.trim();
  page.dataset.pageNumber = String(pageNumber);
  page.dataset.totalPages = String(totalPages);

  const footer = document.createElement("footer");
  footer.className = "page-footer";

  const footerLabel = document.createElement("span");
  footerLabel.className = "page-footer-label";
  footerLabel.textContent = footerText;

  const footerNumber = document.createElement("span");
  footerNumber.className = "page-number";
  footerNumber.textContent = `第 ${pageNumber} / ${totalPages} 頁`;

  footer.append(footerLabel, footerNumber);
  page.appendChild(footer);

  return page;
}

function insertBodyBeforeFooter(page, body) {
  const footer = page.querySelector(".page-footer");
  page.insertBefore(body, footer);
}

function insertHeaderBeforeBody(page, header) {
  const footer = page.querySelector(".page-footer");
  page.insertBefore(header, footer);
}

function renderCompetencyPanel(competencies = []) {
  const panel = document.createElement("section");
  panel.className = "competency-panel";

  const title = document.createElement("h2");
  title.className = "competency-title";
  title.textContent = "核心能力";

  const list = document.createElement("div");
  list.className = "competency-list";

  competencies.forEach((competency) => {
    const item = document.createElement("article");
    item.className = "competency-item";

    const head = document.createElement("div");
    head.className = "competency-head";

    const name = document.createElement("span");
    name.className = "competency-name";
    name.textContent = competency.name;

    const score = document.createElement("span");
    score.className = "competency-score";
    score.textContent = `${competency.score}%`;

    head.append(name, score);

    const track = document.createElement("div");
    track.className = "competency-track";

    const bar = document.createElement("div");
    bar.className = "competency-bar";
    bar.style.width = `${competency.score}%`;

    track.appendChild(bar);
    item.append(head, track);
    list.appendChild(item);
  });

  panel.append(title, list);
  return panel;
}

function renderCoverPortraitPanel(headshot) {
  const panel = document.createElement("section");
  panel.className = "cover-portrait-panel";

  const frame = document.createElement("div");
  frame.className = "cover-photo-frame cover-photo-frame--portrait";

  if (headshot?.src) {
    const img = document.createElement("img");
    img.src = headshot.src;
    img.alt = headshot.alt || "人物照片";
    img.loading = "lazy";
    frame.appendChild(img);
  } else {
    const empty = document.createElement("div");
    empty.className = "cover-photo-empty";
    frame.appendChild(empty);
  }

  panel.appendChild(frame);
  return panel;
}

function renderCoverPage(portfolio, coverImages, totalPages, pageNumber, options = {}) {
  const footerText = `${portfolio.title}｜${portfolio.theme}`;
  const page = createPage(pageNumber, totalPages, footerText, "cover-page");
  page.dataset.exportLabel = "封面";

  if (options.compact) {
    page.classList.add("cover-page--compact");
  }

  const header = document.createElement("header");
  header.className = "page-header";
  header.innerHTML = `
    <div class="page-header-top">
      <span class="page-kicker">${portfolio.theme}</span>
      <span class="page-code">封面</span>
    </div>
  `;

  const body = document.createElement("div");
  body.className = "page-body";

  const hero = document.createElement("section");
  hero.className = "cover-hero";

  const coverMain = document.createElement("div");
  coverMain.className = "cover-main";
  coverMain.innerHTML = `
    <span class="cover-label">${portfolio.theme}</span>
    <h1 class="cover-title">${portfolio.title}</h1>
    <p class="cover-subtitle">
      ${portfolio.subtitleLines.map((line) => `<span class="cover-subtitle-line">${line}</span>`).join("")}
    </p>
    <p class="cover-intro">${portfolio.intro}</p>
  `;

  const coverAside = document.createElement("div");
  coverAside.className = "cover-aside";
  coverAside.append(
    renderCoverPortraitPanel(coverImages.headshot),
    renderCompetencyPanel(portfolio.competencies)
  );
  hero.append(coverMain, coverAside);

  const coverLowerGrid = document.createElement("section");
  coverLowerGrid.className = "cover-lower-grid";

  const outlineCard = document.createElement("section");
  outlineCard.className = "panel-card";

  const outlineTitle = document.createElement("h2");
  outlineTitle.className = "block-title";
  outlineTitle.textContent = "內容目錄";

  const outlineList = document.createElement("ul");
  outlineList.className = "outline-list";

  portfolio.modules.forEach((section) => {
    const item = document.createElement("li");
    item.className = "outline-item";
    item.innerHTML = `
      <span class="outline-index">${getSectionMarker(section.title)}</span>
      <h3 class="outline-name">${getSectionTitleText(section.title)}</h3>
      <span class="outline-count">${section.items.length} 項</span>
    `;
    outlineList.appendChild(item);
  });

  outlineCard.append(outlineTitle, outlineList);

  const highlightPanel = document.createElement("section");
  highlightPanel.className = "highlight-panel";

  const highlightTitle = document.createElement("h2");
  highlightTitle.className = "block-title";
  highlightTitle.textContent = "關鍵亮點";

  const highlightGrid = document.createElement("div");
  highlightGrid.className = "highlight-grid";

  portfolio.highlights.forEach((highlight) => {
    const card = document.createElement("article");
    card.className = "highlight-card";
    card.style.setProperty("--highlight-color", highlight.color);

    if (highlight.material) {
      card.classList.add(`highlight-card--${highlight.material}`);
    }

    card.innerHTML = `
      <p class="highlight-note">${highlight.note}</p>
      <h3 class="highlight-title">${highlight.title}</h3>
    `;

    highlightGrid.appendChild(card);
  });

  highlightPanel.append(highlightTitle, highlightGrid);
  coverLowerGrid.append(outlineCard, highlightPanel);

  body.append(hero, coverLowerGrid);

  insertHeaderBeforeBody(page, header);
  insertBodyBeforeFooter(page, body);

  return page;
}

function renderImageGrid(images, options = {}) {
  const { compact = false } = options;

  if (!images.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "尚未放入佐證圖片。";
    return empty;
  }

  const grid = document.createElement("div");
  grid.className = compact ? "image-grid image-grid--compact" : "image-grid";
  grid.dataset.count = String(images.length);

  images.forEach((image, index) => {
    const frame = document.createElement("figure");
    frame.className = "image-frame";

    const img = document.createElement("img");
    img.src = image.src;
    img.alt = image.alt || `${image.label || "圖片"} ${index + 1}`;
    img.loading = "lazy";

    const label = document.createElement("figcaption");
    label.className = "image-label";
    label.textContent = image.label || `圖片 ${index + 1}`;

    frame.append(img, label);
    grid.appendChild(frame);
  });

  return grid;
}

function renderItemCard(item, itemNumber, section, options = {}) {
  const detailSections = options.detailSections || item.detailSections || [];
  const detailed = detailSections.length > 0;
  const images = options.images || [];
  const showEvidence = options.showEvidence ?? (images.length > 0);
  const descText = options.desc ?? item.desc;
  const evidenceOnly = options.evidenceOnly === true;
  const article = document.createElement("article");
  article.className = detailed ? "item-card item-card--detailed" : "item-card";
  article.style.setProperty("--section-accent", section.accent);
  article.style.setProperty("--section-soft", section.soft);

  const header = document.createElement("div");
  header.className = "item-head";

  const index = document.createElement("span");
  index.className = "item-index";
  index.textContent = `${String(itemNumber).padStart(2, "0")}`;

  const name = document.createElement("h3");
  name.className = "item-name";
  name.textContent = item.name;

  header.append(index, name);

  if (options.continuationText) {
    const continuation = document.createElement("p");
    continuation.className = "item-meta-note";
    continuation.textContent = options.continuationText;
    header.appendChild(continuation);
  }

  const body = document.createElement("div");
  if (evidenceOnly) {
    body.className = "item-body item-body--text-only";
  } else if (detailed) {
    body.className = showEvidence ? "item-body item-body--stacked" : "item-body item-body--stacked item-body--text-only";
  } else {
    body.className = showEvidence ? "item-body" : "item-body item-body--text-only";
  }

  if (!evidenceOnly) {
    const descSection = document.createElement("section");
    descSection.className = "item-panel item-panel--text";

    const descLabel = document.createElement("p");
    descLabel.className = "item-section-label";
    descLabel.textContent = options.descLabel || (detailed ? "內容摘要" : "內容說明");
    descSection.appendChild(descLabel);

    if (descText) {
      const descParagraph = document.createElement("p");
      descParagraph.className = "item-desc";
      descParagraph.textContent = descText;
      descSection.appendChild(descParagraph);
    }

    if (detailed) {
      const detailList = document.createElement("div");
      detailList.className = "item-detail-list";

      detailSections.forEach((sectionDetail) => {
        const block = document.createElement("section");
        block.className = "item-detail-block";

        const heading = document.createElement("h4");
        heading.className = "item-detail-heading";
        heading.textContent = sectionDetail.heading;

        block.appendChild(heading);

        sectionDetail.paragraphs.forEach((paragraph) => {
          const text = document.createElement("p");
          text.className = "item-detail-paragraph";
          text.textContent = paragraph;
          block.appendChild(text);
        });

        detailList.appendChild(block);
      });

      descSection.appendChild(detailList);
    }

    body.appendChild(descSection);
  }

  if (showEvidence) {
    const imageSection = document.createElement("section");
    imageSection.className = "item-panel item-panel--evidence";

    const imageLabel = document.createElement("p");
    imageLabel.className = "item-section-label";
    imageLabel.textContent = "佐證資料";

    imageSection.append(imageLabel, renderImageGrid(images, { compact: detailed }));
    body.appendChild(imageSection);
  }

  article.append(header, body);

  return article;
}

function renderSectionPage(portfolio, section, pageEntry, partIndex, partTotal, totalPages, pageNumber, imageLibrary) {
  const footerText = `${portfolio.title}｜${portfolio.theme}`;
  const page = createPage(pageNumber, totalPages, footerText, "section-page");
  page.style.setProperty("--section-accent", section.accent);
  page.style.setProperty("--section-soft", section.soft);
  page.dataset.exportLabel = `${section.label}｜${getSectionTitleText(section.title)}｜第 ${partIndex} 頁`;

  const header = document.createElement("header");
  header.className = "page-header";
  header.innerHTML = `
    <div class="page-header-top">
      <span class="page-kicker">${portfolio.theme}</span>
      <span class="page-code">${section.label}</span>
    </div>
  `;

  const body = document.createElement("div");
  body.className = "page-body";

  const chapterBand = document.createElement("section");
  chapterBand.className = "chapter-band";
  chapterBand.innerHTML = `
    <span class="chapter-index">${section.title.charAt(0)}</span>
    <span class="chapter-label">${section.label}</span>
    <h2 class="section-title">${section.title}</h2>
    <p class="section-summary">${section.summary}</p>
  `;

  if (partTotal > 1) {
    const continuation = document.createElement("p");
    continuation.className = "continuation-note";
    continuation.textContent = `本部分第 ${partIndex} / ${partTotal} 頁`;
    chapterBand.appendChild(continuation);
  }

  const itemList = document.createElement("div");
  itemList.className = "item-list";

  if (Array.isArray(pageEntry)) {
    pageEntry.forEach((entry) => {
      itemList.appendChild(
        renderItemCard(entry.item, entry.itemNumber, section, {
          images: imageLibrary[entry.item.id] || [],
          showEvidence: true
        })
      );
    });
  } else if (pageEntry.kind === "detail") {
    page.classList.add("section-page--continued");
    page.dataset.itemId = pageEntry.item.id;
    itemList.appendChild(
      renderItemCard(pageEntry.item, pageEntry.itemNumber, section, {
        detailSections: pageEntry.detailPage.detailSections,
        desc: pageEntry.detailPage.desc,
        descLabel: pageEntry.detailPage.descLabel,
        images: pageEntry.detailPage.images,
        showEvidence: pageEntry.detailPage.showEvidence,
        evidenceOnly: pageEntry.detailPage.evidenceOnly,
        continuationText:
          pageEntry.detailPage.pageTotal > 1
            ? `項目第 ${pageEntry.detailPage.pageIndex} / ${pageEntry.detailPage.pageTotal} 頁`
            : ""
      })
    );
  }

  body.append(chapterBand, itemList);

  insertHeaderBeforeBody(page, header);
  insertBodyBeforeFooter(page, body);

  return page;
}

function renderDocument(portfolio, imageData, options = {}) {
  documentRootElement.innerHTML = "";
  currentPortfolio = portfolio;
  const imageLibrary = imageData.images || {};
  const coverImages = imageData.cover || {};
  const detailPageOverrides = options.detailPageOverrides || {};

  const sectionChunks = portfolio.modules.map((section) => ({
    section,
    parts: buildSectionParts(section, imageLibrary, detailPageOverrides)
  }));

  const totalPages = 1 + sectionChunks.reduce((sum, entry) => sum + entry.parts.length, 0);
  let pageNumber = 1;

  documentRootElement.appendChild(
    renderCoverPage(portfolio, coverImages, totalPages, pageNumber, {
      compact: options.coverCompact
    })
  );
  pageNumber += 1;

  sectionChunks.forEach(({ section, parts }) => {
    parts.forEach((pageEntry, partIndex) => {
      documentRootElement.appendChild(
        renderSectionPage(
          portfolio,
          section,
          pageEntry,
          partIndex + 1,
          parts.length,
          totalPages,
          pageNumber,
          imageLibrary
        )
      );

      pageNumber += 1;
    });
  });

  updateExportOptions();
  setExportStatus("可下載整份或指定頁面 PDF");
}

function collectLayoutAdjustmentsFromDiagnostics(diagnostics, adjustments) {
  let changed = false;

  diagnostics.forEach((issue) => {
    const pageElement = documentRootElement.querySelector(`.a4-page[data-page-number="${issue.pageNumber}"]`);

    if (!pageElement) {
      return;
    }

    if (pageElement.classList.contains("cover-page")) {
      if (!adjustments.coverCompact) {
        adjustments.coverCompact = true;
        changed = true;
      }

      return;
    }

    const itemId = pageElement.dataset.itemId;

    if (!itemId) {
      return;
    }

    const nextPageCount = (adjustments.detailPageOverrides[itemId] || 1) + 1;

    if (nextPageCount !== adjustments.detailPageOverrides[itemId]) {
      adjustments.detailPageOverrides[itemId] = nextPageCount;
      changed = true;
    }
  });

  return changed;
}

async function renderDocumentWithOverflowResolution(portfolio, imageData) {
  const adjustments = {
    coverCompact: false,
    detailPageOverrides: {}
  };

  let diagnostics = [];

  for (let iteration = 0; iteration < 8; iteration += 1) {
    renderDocument(portfolio, imageData, adjustments);
    diagnostics = await refreshLayoutDiagnostics();

    const changed = collectLayoutAdjustmentsFromDiagnostics(diagnostics, adjustments);

    if (!changed) {
      return diagnostics;
    }
  }

  renderDocument(portfolio, imageData, adjustments);
  return refreshLayoutDiagnostics();
}

function renderError(message) {
  documentRootElement.innerHTML = "";
  currentPortfolio = null;

  const page = document.createElement("section");
  page.className = "a4-page status-page";

  const card = document.createElement("div");
  card.className = "status-card";
  card.innerHTML = `
    <p class="status-title">內容載入失敗</p>
    <p class="status-text">${message}</p>
    <p class="status-text">請確認資料檔案完整後重新開啟。</p>
  `;

  page.appendChild(card);
  documentRootElement.appendChild(page);
  updateExportOptions();
  renderLayoutDiagnostics([]);
  setExportStatus("內容載入失敗，無法匯出 PDF", "error");
}

function bindExportEvents() {
  downloadAllButtonElement.addEventListener("click", handleExportAllPdf);
  downloadPageButtonElement.addEventListener("click", handleExportSinglePagePdf);
  downloadPageSelectElement.addEventListener("change", refreshExportControls);
  copyLayoutCodeButtonElement?.addEventListener("click", copyDiagnosticCode);
  window.addEventListener("resize", scheduleLayoutDiagnosticRefresh);
  window.addEventListener("load", () => {
    refreshLayoutDiagnostics().catch((error) => console.error(error));
  });
}

async function bootstrap() {
  try {
    const [contentData, imageData] = await Promise.all([
      loadJson("content.json"),
      loadJson("images.json")
    ]);

    await renderDocumentWithOverflowResolution(contentData.portfolio, imageData);
  } catch (error) {
    console.error(error);
    renderError(error.message);
  }
}

bindExportEvents();
refreshExportControls();
bootstrap();
