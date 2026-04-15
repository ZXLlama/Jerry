const PDF_PAGE_WIDTH_PT = 595.28;
const PDF_PAGE_HEIGHT_PT = 841.89;
const EXPORT_RENDER_SCALE = Math.max(2, window.devicePixelRatio || 1);
const A4_RATIO = 297 / 210;
const A4_HEIGHT_TOLERANCE_PX = 6;
const DETAIL_TEXT_CHUNK_SIZE = 180;
const DETAIL_BLOCK_TARGET = 240;
const DETAIL_FIRST_PAGE_CAPACITY = 310;
const DETAIL_CONTINUED_PAGE_CAPACITY = 440;
const PAGE_IMAGE_LIMIT = 2;

const documentRootElement = document.getElementById("document-root");
const downloadAllButtonElement = document.getElementById("download-all-pdf");
const downloadPageButtonElement = document.getElementById("download-page-pdf");
const downloadPageSelectElement = document.getElementById("download-page-select");
const exportStatusElement = document.getElementById("export-status");
const diagnosticPanelElement = document.getElementById("diagnostic-panel");
const diagnosticSummaryElement = document.getElementById("diagnostic-summary");
const diagnosticCodeElement = document.getElementById("diagnostic-code");
const copyLayoutCodeButtonElement = document.getElementById("copy-layout-code");

let currentPortfolio = null;
let currentImageData = null;
let currentLayoutDiagnostics = [];
let exportInProgress = false;
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

function getDocumentFileName() {
  return `${getDocumentBaseName()}_完整.pdf`;
}

function getPageExportLabel(pageElement, pageIndex) {
  return pageElement.dataset.exportLabel?.trim() || `第 ${pageIndex + 1} 頁`;
}

function getPageFileName(pageElement, pageIndex) {
  const label = sanitizeFileName(getPageExportLabel(pageElement, pageIndex));
  const pageNumber = String(pageIndex + 1).padStart(2, "0");
  return `${getDocumentBaseName()}_頁面_${pageNumber}_${label}.pdf`;
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
    exportInProgress || !hasPages || !downloadPageSelectElement.value;
}

function updateExportOptions() {
  const pages = getDocumentPages();
  const previousValue = downloadPageSelectElement.value;
  downloadPageSelectElement.innerHTML = "";

  if (!pages.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "請先載入內容";
    downloadPageSelectElement.appendChild(option);
    refreshExportControls();
    return;
  }

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "選擇單頁 PDF";
  downloadPageSelectElement.appendChild(placeholder);

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

function splitParagraphIntoChunks(paragraph, chunkSize = DETAIL_TEXT_CHUNK_SIZE) {
  const normalized = String(paragraph || "").trim();

  if (!normalized) {
    return [];
  }

  const sentenceParts =
    normalized.match(/[^。！？]+[。！？]?/gu)?.map((part) => part.trim()).filter(Boolean) || [
      normalized
    ];

  const chunks = [];
  let current = "";

  sentenceParts.forEach((part) => {
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
    if (chunk.length <= chunkSize * 1.25) {
      return [chunk];
    }

    const refinedParts =
      chunk.match(/[^，；]+[，；]?/gu)?.map((part) => part.trim()).filter(Boolean) || [chunk];
    const refined = [];
    let refinedCurrent = "";

    refinedParts.forEach((part) => {
      if (!refinedCurrent || refinedCurrent.length + part.length <= chunkSize) {
        refinedCurrent += part;
        return;
      }

      refined.push(refinedCurrent);
      refinedCurrent = part;
    });

    if (refinedCurrent) {
      refined.push(refinedCurrent);
    }

    return refined;
  });
}

function createDetailBlocks(item) {
  return (item.detailSections || []).flatMap((detailSection) => {
    const paragraphChunks = (detailSection.paragraphs || []).flatMap((paragraph) =>
      splitParagraphIntoChunks(paragraph)
    );

    const blocks = [];
    let currentParagraphs = [];
    let currentLength = 0;

    const pushBlock = () => {
      if (!currentParagraphs.length) {
        return;
      }

      const blockIndex = blocks.length;
      blocks.push({
        heading:
          blockIndex === 0 ? detailSection.heading : `${detailSection.heading}（續）`,
        paragraphs: [...currentParagraphs]
      });

      currentParagraphs = [];
      currentLength = 0;
    };

    paragraphChunks.forEach((chunk) => {
      if (currentParagraphs.length && currentLength + chunk.length > DETAIL_BLOCK_TARGET) {
        pushBlock();
      }

      currentParagraphs.push(chunk);
      currentLength += chunk.length;
    });

    pushBlock();
    return blocks;
  });
}

function getDetailBlockWeight(block) {
  return (block.heading || "").length * 4 + (block.paragraphs || []).join("").length;
}

function distributeDetailBlocks(blocks, pageCount, imageCount) {
  const pages = Array.from({ length: pageCount }, () => []);

  if (!blocks.length) {
    return pages;
  }

  const capacities = Array.from({ length: pageCount }, (_, index) => {
    if (index === 0) {
      return Math.max(170, DETAIL_FIRST_PAGE_CAPACITY - imageCount * 26);
    }

    return DETAIL_CONTINUED_PAGE_CAPACITY;
  });

  let pageIndex = 0;
  let pageWeight = 0;

  blocks.forEach((block, blockIndex) => {
    const weight = getDetailBlockWeight(block);
    const remainingBlocks = blocks.length - blockIndex;
    const remainingPages = pageCount - pageIndex;
    const shouldAdvance =
      pageIndex < pageCount - 1 &&
      pages[pageIndex].length > 0 &&
      (pageWeight + weight > capacities[pageIndex] || remainingBlocks === remainingPages);

    if (shouldAdvance) {
      pageIndex += 1;
      pageWeight = 0;
    }

    pages[pageIndex].push(block);
    pageWeight += weight;
  });

  return pages;
}

function distributeImages(images, pageCount) {
  const pages = Array.from({ length: pageCount }, () => []);
  let cursor = 0;

  for (let round = 0; round < PAGE_IMAGE_LIMIT; round += 1) {
    for (let pageIndex = 0; pageIndex < pageCount && cursor < images.length; pageIndex += 1) {
      pages[pageIndex].push(images[cursor]);
      cursor += 1;
    }
  }

  while (cursor < images.length) {
    pages[pageCount - 1].push(images[cursor]);
    cursor += 1;
  }

  return pages;
}

function buildItemPages(item, imageLibrary, requestedPageCount = 1) {
  const images = imageLibrary[item.id] || [];
  const detailBlocks = createDetailBlocks(item);
  const minimumPageCount = Math.max(1, Math.ceil(images.length / PAGE_IMAGE_LIMIT));
  const targetPageCount = Math.max(requestedPageCount, minimumPageCount);
  const blocksByPage = distributeDetailBlocks(detailBlocks, targetPageCount, images.length);
  const imagesByPage = distributeImages(images, targetPageCount);
  const rawPages = [];

  for (let index = 0; index < targetPageCount; index += 1) {
    const pageBlocks = blocksByPage[index] || [];
    const pageImages = imagesByPage[index] || [];

    if (!pageBlocks.length && !pageImages.length) {
      continue;
    }

    rawPages.push({
      item,
      detailBlocks: pageBlocks,
      images: pageImages
    });
  }

  return rawPages.map((page, index) => ({
    ...page,
    pageIndex: index + 1,
    pageTotal: rawPages.length
  }));
}

function renderText(elementTag, className, text) {
  const element = document.createElement(elementTag);
  element.className = className;
  element.textContent = text;
  return element;
}

function renderModuleStrip(modules, activeModuleId = "") {
  const strip = document.createElement("nav");
  strip.className = "module-strip";
  strip.setAttribute("aria-label", "主題導覽");

  modules.forEach((module, index) => {
    const item = document.createElement("div");
    item.className = "module-strip__item";
    item.style.setProperty("--module-accent", module.accent);

    if (module.id === activeModuleId) {
      item.classList.add("is-active");
    }

    const number = renderText("span", "module-strip__number", String(index + 1));
    const title = renderText("span", "module-strip__title", module.navTitle);
    item.append(number, title);
    strip.appendChild(item);
  });

  return strip;
}

function renderDashboard(portfolio) {
  const panel = document.createElement("section");
  panel.className = "dashboard-panel";

  panel.appendChild(renderText("h2", "panel-title", portfolio.dashboardTitle));

  const list = document.createElement("div");
  list.className = "dashboard-list";

  (portfolio.dashboard || []).forEach((item) => {
    const row = document.createElement("article");
    row.className = "dashboard-row";
    row.dataset.tone = item.tone;

    const label = renderText("p", "dashboard-row__label", item.name);
    const track = document.createElement("div");
    track.className = "dashboard-row__track";

    const fill = document.createElement("div");
    fill.className = "dashboard-row__fill";
    fill.style.width = `${Math.round(item.strength * 100)}%`;

    const cap = document.createElement("span");
    cap.className = "dashboard-row__cap";

    track.append(fill, cap);
    row.append(label, track);
    list.appendChild(row);
  });

  panel.appendChild(list);
  return panel;
}

function renderHighlightPanel(portfolio) {
  const panel = document.createElement("section");
  panel.className = "highlight-panel";
  panel.appendChild(renderText("h2", "panel-title", portfolio.highlightsTitle));

  const grid = document.createElement("div");
  grid.className = "highlight-grid";

  (portfolio.highlights || []).forEach((highlight) => {
    const card = document.createElement("article");
    card.className = "highlight-card";
    card.dataset.tone = highlight.tone;

    card.append(
      renderText("p", "highlight-card__label", highlight.label),
      renderText("h3", "highlight-card__value", highlight.value)
    );

    grid.appendChild(card);
  });

  panel.appendChild(grid);
  return panel;
}

function renderOutlinePanel(portfolio) {
  const panel = document.createElement("section");
  panel.className = "outline-panel";
  panel.appendChild(renderText("h2", "panel-title", "內容目錄"));

  const list = document.createElement("div");
  list.className = "outline-list";

  (portfolio.modules || []).forEach((module, index) => {
    const row = document.createElement("article");
    row.className = "outline-row";
    row.style.setProperty("--module-accent", module.accent);

    const indexWrap = document.createElement("div");
    indexWrap.className = "outline-row__index";
    indexWrap.textContent = `0${index + 1}`.slice(-2);

    const content = document.createElement("div");
    content.className = "outline-row__content";
    content.append(
      renderText("h3", "outline-row__title", module.title),
      renderText("p", "outline-row__subtitle", module.subtitle)
    );

    const count = renderText("div", "outline-row__count", `${module.items.length} 項`);

    row.append(indexWrap, content, count);
    list.appendChild(row);
  });

  panel.appendChild(list);
  return panel;
}

function renderCoverPortrait(coverImages) {
  const panel = document.createElement("section");
  panel.className = "cover-portrait-panel";

  const frame = document.createElement("div");
  frame.className = "cover-portrait-frame";

  if (coverImages?.headshot?.src) {
    const image = document.createElement("img");
    image.src = coverImages.headshot.src;
    image.alt = coverImages.headshot.alt || "黃煥傑個人照片";
    image.loading = "lazy";
    frame.appendChild(image);
  }

  panel.appendChild(frame);
  return panel;
}

function createPageShell(pageNumber, totalPages, footerText, pageClass = "") {
  const page = document.createElement("section");
  page.className = `a4-page ${pageClass}`.trim();
  page.dataset.pageNumber = String(pageNumber);
  page.dataset.totalPages = String(totalPages);

  const header = document.createElement("header");
  header.className = "page-header";

  const body = document.createElement("div");
  body.className = "page-body";

  const footer = document.createElement("footer");
  footer.className = "page-footer";
  footer.append(
    renderText("span", "page-footer__label", footerText),
    renderText("span", "page-footer__number", `第 ${pageNumber} / ${totalPages} 頁`)
  );

  page.append(header, body, footer);
  return { page, header, body };
}

function populatePageHeader(header, leftText, rightText) {
  header.innerHTML = "";

  const left = renderText("span", "page-header__left", leftText);
  const right = renderText("span", "page-header__right page-code", rightText);
  header.append(left, right);
}

function renderCoverPage(portfolio, coverImages, totalPages, pageNumber, options = {}) {
  const shell = createPageShell(
    pageNumber,
    totalPages,
    `${portfolio.title}｜${portfolio.theme}`,
    "cover-page"
  );

  shell.page.dataset.exportLabel = "封面";

  if (options.compact) {
    shell.page.classList.add("cover-page--compact");
  }

  populatePageHeader(shell.header, portfolio.theme, "封面");
  shell.body.appendChild(renderModuleStrip(portfolio.modules));

  const hero = document.createElement("section");
  hero.className = "cover-hero";

  const intro = document.createElement("section");
  intro.className = "cover-intro";
  intro.append(
    renderText("p", "cover-note", portfolio.coverNote),
    renderText("p", "cover-theme", portfolio.theme),
    renderText("h1", "cover-title", portfolio.title)
  );

  const school = document.createElement("div");
  school.className = "cover-school";
  (portfolio.schoolLines || []).forEach((line) => {
    school.appendChild(renderText("p", "cover-school__line", line));
  });

  intro.append(school, renderText("p", "cover-tagline", portfolio.coverTagline));

  const rightColumn = document.createElement("div");
  rightColumn.className = "cover-side";
  rightColumn.append(renderCoverPortrait(coverImages), renderDashboard(portfolio));

  hero.append(intro, rightColumn);

  const lower = document.createElement("section");
  lower.className = "cover-lower";
  lower.append(renderOutlinePanel(portfolio), renderHighlightPanel(portfolio));

  shell.body.append(hero, lower);
  return shell.page;
}

function renderStoryHero(module, itemPage) {
  const hero = document.createElement("section");
  hero.className = "story-hero";
  hero.style.setProperty("--module-accent", module.accent);

  const meta = document.createElement("div");
  meta.className = "story-hero__meta";
  meta.append(
    renderText("p", "story-hero__part", `${module.partLabel}｜${module.title}`),
    renderText("p", "story-hero__subtitle", module.subtitle)
  );

  const serial = renderText("div", "story-hero__serial", itemPage.item.serial);
  const content = document.createElement("div");
  content.className = "story-hero__content";
  content.append(
    renderText("p", "story-hero__eyebrow", itemPage.item.eyebrow),
    renderText("h2", "story-hero__title", itemPage.item.title)
  );

  hero.append(meta, serial, content);

  if (itemPage.pageTotal > 1) {
    hero.appendChild(
      renderText("p", "story-hero__note", `續頁 ${itemPage.pageIndex} / ${itemPage.pageTotal}`)
    );
  }

  return hero;
}

function renderImageFrame(image, modifier = "") {
  const frame = document.createElement("figure");
  frame.className = `image-frame ${modifier}`.trim();

  const picture = document.createElement("img");
  picture.src = image.src;
  picture.alt = image.alt || image.label || "佐證圖片";
  picture.loading = "lazy";

  const caption = renderText("figcaption", "image-frame__caption", image.label || "證據");
  frame.append(picture, caption);
  return frame;
}

function renderEvidenceGallery(images, title, prominent = false) {
  const section = document.createElement("section");
  section.className = "evidence-gallery";

  section.appendChild(renderText("p", "section-eyebrow", title));

  const grid = document.createElement("div");
  grid.className = prominent ? "evidence-gallery__grid is-prominent" : "evidence-gallery__grid";

  images.forEach((image) => {
    grid.appendChild(renderImageFrame(image, prominent ? "is-large" : ""));
  });

  section.appendChild(grid);
  return section;
}

function renderStoryRows(detailBlocks, images) {
  const section = document.createElement("section");
  section.className = "story-list";

  if (!detailBlocks.length) {
    section.classList.add("story-list--evidence-only");
    section.appendChild(renderEvidenceGallery(images, "佐證圖片", true));
    return section;
  }

  let imageCursor = 0;

  detailBlocks.forEach((block) => {
    const row = document.createElement("article");
    row.className = "story-row";

    const text = document.createElement("div");
    text.className = "story-row__text";
    text.appendChild(renderText("h3", "story-row__heading", block.heading));

    (block.paragraphs || []).forEach((paragraph) => {
      text.appendChild(renderText("p", "story-row__paragraph", paragraph));
    });

    row.appendChild(text);

    if (images[imageCursor]) {
      const figureWrap = document.createElement("div");
      figureWrap.className = "story-row__figure";
      figureWrap.appendChild(renderImageFrame(images[imageCursor]));
      row.appendChild(figureWrap);
      imageCursor += 1;
    } else {
      row.classList.add("is-text-only");
    }

    section.appendChild(row);
  });

  const remainingImages = images.slice(imageCursor);

  if (remainingImages.length) {
    section.appendChild(renderEvidenceGallery(remainingImages, "補充佐證"));
  }

  return section;
}

function renderStoryPage(portfolio, module, itemPage, totalPages, pageNumber) {
  const shell = createPageShell(
    pageNumber,
    totalPages,
    `${portfolio.title}｜${portfolio.theme}`,
    "story-page"
  );

  shell.page.dataset.itemId = itemPage.item.id;
  shell.page.dataset.moduleId = module.id;
  shell.page.dataset.exportLabel =
    `${module.partLabel}｜${itemPage.item.title}｜第 ${itemPage.pageIndex} 頁`;

  if (itemPage.pageIndex > 1) {
    shell.page.classList.add("story-page--continuation");
  }

  if (!itemPage.detailBlocks.length) {
    shell.page.classList.add("story-page--evidence");
  }

  populatePageHeader(shell.header, portfolio.theme, itemPage.item.serial);
  shell.body.append(
    renderModuleStrip(portfolio.modules, module.id),
    renderStoryHero(module, itemPage),
    renderStoryRows(itemPage.detailBlocks, itemPage.images)
  );

  return shell.page;
}

function renderDocument(portfolio, imageData, options = {}) {
  currentPortfolio = portfolio;
  currentImageData = imageData;
  documentRootElement.innerHTML = "";

  const itemPageOverrides = options.itemPageOverrides || {};
  const modules = (portfolio.modules || []).map((module) => ({
    module,
    pages: module.items.flatMap((item) =>
      buildItemPages(item, imageData.images || {}, itemPageOverrides[item.id] || 1)
    )
  }));

  const totalPages = 1 + modules.reduce((sum, entry) => sum + entry.pages.length, 0);
  let pageNumber = 1;

  documentRootElement.appendChild(
    renderCoverPage(portfolio, imageData.cover || {}, totalPages, pageNumber, {
      compact: options.coverCompact
    })
  );
  pageNumber += 1;

  modules.forEach(({ module, pages }) => {
    pages.forEach((itemPage) => {
      documentRootElement.appendChild(
        renderStoryPage(portfolio, module, itemPage, totalPages, pageNumber)
      );
      pageNumber += 1;
    });
  });

  updateExportOptions();
  setExportStatus("內容已更新，可下載 PDF。");
}

function getExpectedPageHeightPx(pageElement) {
  return pageElement.getBoundingClientRect().width * A4_RATIO;
}

function getPageItemLabels(pageElement) {
  if (pageElement.classList.contains("cover-page")) {
    return ["封面"];
  }

  const title = pageElement.querySelector(".story-hero__title")?.textContent?.trim();
  return title ? [title] : [pageElement.dataset.exportLabel || "內容頁"];
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

      return {
        pageNumber,
        pageLabel: getPageExportLabel(pageElement, pageNumber - 1),
        itemLabels: getPageItemLabels(pageElement),
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

  const pageMap = new Map(
    getDocumentPages().map((pageElement) => [Number(pageElement.dataset.pageNumber), pageElement])
  );

  diagnostics.forEach((diagnostic) => {
    const pageElement = pageMap.get(diagnostic.pageNumber);

    if (!pageElement) {
      return;
    }

    pageElement.dataset.layoutWarning = "true";
    pageElement.dataset.warningCode = diagnostic.code;

    const warning = document.createElement("aside");
    warning.className = "page-layout-warning pdf-control";
    warning.innerHTML = `
      <span class="page-layout-warning__label">PDF 匯出風險</span>
      <strong class="page-layout-warning__code">${diagnostic.code}</strong>
      <span class="page-layout-warning__meta">超出 A4 ${diagnostic.overflowPx}px</span>
    `;

    pageElement.appendChild(warning);
  });
}

function renderLayoutDiagnostics(diagnostics) {
  currentLayoutDiagnostics = diagnostics;
  renderPageLayoutWarnings(diagnostics);

  if (!diagnosticPanelElement) {
    return;
  }

  if (!diagnostics.length) {
    diagnosticPanelElement.hidden = true;
    diagnosticSummaryElement.textContent = "";
    diagnosticCodeElement.textContent = "";
    copyLayoutCodeButtonElement.disabled = true;
    return;
  }

  const pageList = diagnostics.map((issue) => issue.pageNumber).join("、");
  diagnosticPanelElement.hidden = false;
  copyLayoutCodeButtonElement.disabled = false;
  diagnosticSummaryElement.textContent =
    `第 ${pageList} 頁超出 A4 高度，若直接輸出 PDF 將造成整頁被垂直壓縮。請複製以下錯誤代碼交由 Codex 解析。`;
  diagnosticCodeElement.textContent = buildLayoutDiagnosticPayload(diagnostics);
}

async function waitForLayoutStability() {
  if (document.fonts?.ready) {
    await document.fonts.ready;
  }

  await new Promise((resolve) =>
    window.requestAnimationFrame(() => window.requestAnimationFrame(resolve))
  );
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

  return diagnostics.filter((diagnostic) => pageNumbers.has(diagnostic.pageNumber));
}

function focusDiagnosticPanel() {
  if (!diagnosticPanelElement.hidden) {
    diagnosticPanelElement.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }

  documentRootElement
    .querySelector('.a4-page[data-layout-warning="true"]')
    ?.scrollIntoView({ behavior: "smooth", block: "center" });
}

async function copyDiagnosticCode() {
  const text = diagnosticCodeElement.textContent?.trim();

  if (!text) {
    setExportStatus("目前沒有可複製的錯誤代碼。");
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

    setExportStatus("錯誤代碼已複製。", "success");
  } catch (error) {
    console.error(error);
    setExportStatus("複製失敗，請手動複製錯誤代碼。", "error");
  }
}

function scheduleLayoutDiagnosticRefresh() {
  window.clearTimeout(resizeDiagnosticTimer);
  resizeDiagnosticTimer = window.setTimeout(() => {
    refreshLayoutDiagnostics().catch((error) => console.error(error));
  }, 180);
}

function collectLayoutAdjustmentsFromDiagnostics(diagnostics, adjustments) {
  let changed = false;

  diagnostics.forEach((issue) => {
    const pageElement = documentRootElement.querySelector(
      `.a4-page[data-page-number="${issue.pageNumber}"]`
    );

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

    const nextCount = (adjustments.itemPageOverrides[itemId] || 1) + 1;

    if (adjustments.itemPageOverrides[itemId] !== nextCount) {
      adjustments.itemPageOverrides[itemId] = nextCount;
      changed = true;
    }
  });

  return changed;
}

async function renderDocumentWithOverflowResolution(portfolio, imageData) {
  const adjustments = {
    coverCompact: false,
    itemPageOverrides: {}
  };

  for (let iteration = 0; iteration < 10; iteration += 1) {
    renderDocument(portfolio, imageData, adjustments);
    const diagnostics = await refreshLayoutDiagnostics();

    if (!collectLayoutAdjustmentsFromDiagnostics(diagnostics, adjustments)) {
      return diagnostics;
    }
  }

  renderDocument(portfolio, imageData, adjustments);
  return refreshLayoutDiagnostics();
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
    throw new Error("PDF 匯出功能載入失敗。");
  }

  const rect = pageElement.getBoundingClientRect();
  const width = Math.ceil(rect.width);
  const height = Math.ceil(rect.height);

  return window.html2canvas(pageElement, {
    backgroundColor: "#FFFFFF",
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

      const clonedPage = clonedDocument.querySelector(
        `.a4-page[data-page-number="${pageElement.dataset.pageNumber}"]`
      );

      if (!clonedPage) {
        return;
      }

      clonedPage.style.width = `${width}px`;
      clonedPage.style.minHeight = `${height}px`;
      clonedPage.style.height = `${height}px`;
      clonedPage.style.maxHeight = `${height}px`;
      clonedPage.style.margin = "0";
      clonedPage.style.transform = "none";
    }
  });
}

function createPdfDocument() {
  const jsPDFConstructor = window.jspdf?.jsPDF;

  if (!jsPDFConstructor) {
    throw new Error("PDF 匯出功能載入失敗。");
  }

  return new jsPDFConstructor({
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

  const imageData = canvas.toDataURL("image/png", 1);
  pdf.addImage(imageData, "PNG", 0, 0, PDF_PAGE_WIDTH_PT, PDF_PAGE_HEIGHT_PT, undefined, "FAST");
}

async function exportPagesToPdf(pageElements, filename, label) {
  const diagnostics = await refreshLayoutDiagnostics();
  const blockingIssues = getDiagnosticsForPages(pageElements, diagnostics);

  if (blockingIssues.length) {
    const pages = blockingIssues.map((issue) => issue.pageNumber).join("、");
    setExportStatus(
      `${label}已停止。第 ${pages} 頁超出 A4 高度，直接輸出會造成版面被壓縮。`,
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
      setExportStatus(`${label}輸出中 ${index + 1} / ${pageElements.length}`);
      const canvas = await renderPageToCanvas(pageElements[index]);
      appendCanvasToPdf(pdf, canvas, index);
      canvas.width = 0;
      canvas.height = 0;
    }

    triggerFileDownload(pdf.output("blob"), filename);
    setExportStatus(`${label}已完成。`, "success");
  } catch (error) {
    console.error(error);
    setExportStatus(error.message || "PDF 匯出失敗。", "error");
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
  const page = pages[pageIndex];

  if (!page) {
    return;
  }

  exportPagesToPdf([page], getPageFileName(page, pageIndex), `第 ${pageIndex + 1} 頁 PDF`);
}

function renderError(message) {
  currentPortfolio = null;
  currentImageData = null;
  documentRootElement.innerHTML = "";

  const shell = createPageShell(1, 1, "多元表現", "status-page");
  populatePageHeader(shell.header, "多元表現", "狀態");

  const card = document.createElement("div");
  card.className = "status-card";
  card.append(
    renderText("p", "status-card__title", "內容載入失敗"),
    renderText("p", "status-card__text", message),
    renderText("p", "status-card__text", "請確認 JSON 與圖片路徑是否正確。")
  );

  shell.body.appendChild(card);
  documentRootElement.appendChild(shell.page);
  updateExportOptions();
  renderLayoutDiagnostics([]);
  setExportStatus("內容載入失敗，無法輸出 PDF。", "error");
}

function bindExportEvents() {
  downloadAllButtonElement.addEventListener("click", handleExportAllPdf);
  downloadPageButtonElement.addEventListener("click", handleExportSinglePagePdf);
  downloadPageSelectElement.addEventListener("change", refreshExportControls);
  copyLayoutCodeButtonElement?.addEventListener("click", copyDiagnosticCode);
  window.addEventListener("resize", scheduleLayoutDiagnosticRefresh);
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
