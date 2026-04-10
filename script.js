const ITEMS_PER_PAGE = 2;
const PDF_PAGE_WIDTH_PT = 595.28;
const PDF_PAGE_HEIGHT_PT = 841.89;
const EXPORT_RENDER_SCALE = Math.max(2, window.devicePixelRatio || 1);

const documentRootElement = document.getElementById("document-root");
const downloadAllButtonElement = document.getElementById("download-all-pdf");
const downloadPageButtonElement = document.getElementById("download-page-pdf");
const downloadPageSelectElement = document.getElementById("download-page-select");
const exportStatusElement = document.getElementById("export-status");

let exportInProgress = false;
let currentPortfolio = null;

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

function buildSectionParts(section) {
  const parts = [];
  let briefItems = [];

  section.items.forEach((item) => {
    if (hasDetailSections(item)) {
      if (briefItems.length) {
        parts.push(briefItems);
        briefItems = [];
      }

      parts.push([item]);
      return;
    }

    briefItems.push(item);

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

function renderCoverPage(portfolio, coverImages, totalPages, pageNumber) {
  const footerText = `${portfolio.title}｜${portfolio.theme}`;
  const page = createPage(pageNumber, totalPages, footerText, "cover-page");
  page.dataset.exportLabel = "封面";

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

function renderItemCard(item, itemNumber, imageLibrary, section) {
  const detailed = hasDetailSections(item);
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

  const body = document.createElement("div");
  body.className = detailed ? "item-body item-body--stacked" : "item-body";

  const descSection = document.createElement("section");
  descSection.className = "item-panel item-panel--text";
  descSection.innerHTML = `
    <p class="item-section-label">${detailed ? "內容摘要" : "內容說明"}</p>
    <p class="item-desc">${item.desc}</p>
  `;

  if (detailed) {
    const detailList = document.createElement("div");
    detailList.className = "item-detail-list";

    item.detailSections.forEach((sectionDetail) => {
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

  const imageSection = document.createElement("section");
  imageSection.className = "item-panel item-panel--evidence";

  const imageLabel = document.createElement("p");
  imageLabel.className = "item-section-label";
  imageLabel.textContent = "佐證資料";

  imageSection.append(imageLabel, renderImageGrid(imageLibrary[item.id] || [], { compact: detailed }));
  body.append(descSection, imageSection);
  article.append(header, body);

  return article;
}

function renderSectionPage(portfolio, section, items, partIndex, partTotal, totalPages, pageNumber, itemOffset, imageLibrary) {
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

  items.forEach((item, index) => {
    itemList.appendChild(renderItemCard(item, itemOffset + index + 1, imageLibrary, section));
  });

  body.append(chapterBand, itemList);

  insertHeaderBeforeBody(page, header);
  insertBodyBeforeFooter(page, body);

  return page;
}

function renderDocument(portfolio, imageData) {
  documentRootElement.innerHTML = "";
  currentPortfolio = portfolio;
  const imageLibrary = imageData.images || {};
  const coverImages = imageData.cover || {};

  const sectionChunks = portfolio.modules.map((section) => ({
    section,
    parts: buildSectionParts(section)
  }));

  const totalPages = 1 + sectionChunks.reduce((sum, entry) => sum + entry.parts.length, 0);
  let pageNumber = 1;

  documentRootElement.appendChild(renderCoverPage(portfolio, coverImages, totalPages, pageNumber));
  pageNumber += 1;

  sectionChunks.forEach(({ section, parts }) => {
    let itemOffset = 0;

    parts.forEach((items, partIndex) => {
      documentRootElement.appendChild(
        renderSectionPage(
          portfolio,
          section,
          items,
          partIndex + 1,
          parts.length,
          totalPages,
          pageNumber,
          itemOffset,
          imageLibrary
        )
      );

      itemOffset += items.length;
      pageNumber += 1;
    });
  });

  updateExportOptions();
  setExportStatus("可下載整份或指定頁面 PDF");
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
  setExportStatus("內容載入失敗，無法匯出 PDF", "error");
}

function bindExportEvents() {
  downloadAllButtonElement.addEventListener("click", handleExportAllPdf);
  downloadPageButtonElement.addEventListener("click", handleExportSinglePagePdf);
  downloadPageSelectElement.addEventListener("change", refreshExportControls);
}

async function bootstrap() {
  try {
    const [contentData, imageData] = await Promise.all([
      loadJson("content.json"),
      loadJson("images.json")
    ]);

    renderDocument(contentData.portfolio, imageData);
  } catch (error) {
    console.error(error);
    renderError(error.message);
  }
}

bindExportEvents();
refreshExportControls();
bootstrap();
