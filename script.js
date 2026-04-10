const ITEMS_PER_PAGE = 2;
const PDF_PAGE_WIDTH_PT = 595.28;
const PDF_PAGE_HEIGHT_PT = 841.89;
const EXPORT_RENDER_SCALE = Math.max(2, window.devicePixelRatio || 1);

const documentRootElement = document.getElementById("document-root");
const downloadAllButtonElement = document.getElementById("download-all-pdf");
const downloadPageButtonElement = document.getElementById("download-page-pdf");
const downloadPageSelectElement = document.getElementById("download-page-select");
const exportStatusElement = document.getElementById("export-status");

const textEncoder = new TextEncoder();
const resourceDataUrlCache = new Map();

let exportStylesheetTextPromise;
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

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("圖片轉換失敗"));
    reader.readAsDataURL(blob);
  });
}

async function getResourceDataUrl(url) {
  if (!url) {
    return "";
  }

  if (url.startsWith("data:")) {
    return url;
  }

  const absoluteUrl = new URL(url, window.location.href).href;

  if (!resourceDataUrlCache.has(absoluteUrl)) {
    const pending = fetch(absoluteUrl, { cache: "force-cache" })
      .then((response) => {
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        return response.blob();
      })
      .then((blob) => blobToDataUrl(blob));

    resourceDataUrlCache.set(absoluteUrl, pending);
  }

  return resourceDataUrlCache.get(absoluteUrl);
}

function imageElementToDataUrl(imageElement) {
  const canvas = document.createElement("canvas");
  canvas.width = imageElement.naturalWidth;
  canvas.height = imageElement.naturalHeight;

  const context = canvas.getContext("2d");
  context.drawImage(imageElement, 0, 0);

  return canvas.toDataURL("image/png");
}

async function inlineCloneImages(clone, sourcePage) {
  const cloneImages = Array.from(clone.querySelectorAll("img"));
  const sourceImages = Array.from(sourcePage.querySelectorAll("img"));

  await Promise.all(cloneImages.map(async (image, index) => {
    const sourceImage = sourceImages[index];
    const sourceUrl = sourceImage?.currentSrc || sourceImage?.src || image.getAttribute("src");

    if (!sourceUrl) {
      return;
    }

    try {
      image.src = await getResourceDataUrl(sourceUrl);
    } catch (error) {
      try {
        image.src = imageElementToDataUrl(sourceImage);
      } catch (fallbackError) {
        image.src = sourceUrl;
      }
    }

    image.removeAttribute("loading");
    image.removeAttribute("srcset");
    image.setAttribute("decoding", "sync");
  }));
}

async function getExportStylesheetText() {
  if (!exportStylesheetTextPromise) {
    exportStylesheetTextPromise = Promise.resolve(
      Array.from(document.styleSheets)
        .map((sheet) => {
          try {
            return Array.from(sheet.cssRules)
              .map((rule) => rule.cssText)
              .join("\n");
          } catch (error) {
            return "";
          }
        })
        .filter(Boolean)
        .join("\n")
    );
  }

  return exportStylesheetTextPromise;
}

function loadImageElement(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.decoding = "sync";
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("頁面匯出失敗"));
    image.src = src;
  });
}

function concatUint8Arrays(parts) {
  const totalLength = parts.reduce((sum, part) => sum + part.length, 0);
  const merged = new Uint8Array(totalLength);
  let offset = 0;

  parts.forEach((part) => {
    merged.set(part, offset);
    offset += part.length;
  });

  return merged;
}

async function renderPageToCanvas(pageElement) {
  await waitForPageAssets(pageElement);

  const rect = pageElement.getBoundingClientRect();
  const width = Math.ceil(rect.width);
  const height = Math.ceil(rect.height);
  const clone = pageElement.cloneNode(true);

  clone.querySelectorAll(".pdf-control").forEach((node) => node.remove());
  clone.style.width = `${width}px`;
  clone.style.minWidth = `${width}px`;
  clone.style.minHeight = `${height}px`;
  clone.style.margin = "0";
  clone.style.boxShadow = "none";

  await inlineCloneImages(clone, pageElement);

  const stylesheetText = await getExportStylesheetText();
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("width", String(width));
  svg.setAttribute("height", String(height));
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);

  const foreignObject = document.createElementNS("http://www.w3.org/2000/svg", "foreignObject");
  foreignObject.setAttribute("width", "100%");
  foreignObject.setAttribute("height", "100%");

  const shell = document.createElement("div");
  shell.setAttribute("xmlns", "http://www.w3.org/1999/xhtml");
  shell.style.width = `${width}px`;
  shell.style.height = `${height}px`;
  shell.style.background = "#fff";

  const style = document.createElement("style");
  style.textContent = `
    ${stylesheetText}
    .pdf-control { display: none !important; }
    .document-root { padding: 0 !important; gap: 0 !important; }
    body { margin: 0 !important; background: #fff !important; }
    .a4-page {
      width: ${width}px !important;
      min-width: ${width}px !important;
      min-height: ${height}px !important;
      margin: 0 !important;
      box-shadow: none !important;
    }
  `;

  shell.append(style, clone);
  foreignObject.appendChild(shell);
  svg.appendChild(foreignObject);

  const serialized = new XMLSerializer().serializeToString(svg);
  const svgBlob = new Blob([serialized], { type: "image/svg+xml;charset=utf-8" });
  const svgUrl = URL.createObjectURL(svgBlob);

  try {
    const image = await loadImageElement(svgUrl);
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(width * EXPORT_RENDER_SCALE);
    canvas.height = Math.round(height * EXPORT_RENDER_SCALE);

    const context = canvas.getContext("2d");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.setTransform(EXPORT_RENDER_SCALE, 0, 0, EXPORT_RENDER_SCALE, 0, 0);
    context.drawImage(image, 0, 0, width, height);

    return canvas;
  } finally {
    URL.revokeObjectURL(svgUrl);
  }
}

function canvasToJpegBytes(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(async (blob) => {
      if (!blob) {
        reject(new Error("PDF 轉換失敗"));
        return;
      }

      const buffer = await blob.arrayBuffer();
      resolve({
        bytes: new Uint8Array(buffer),
        width: canvas.width,
        height: canvas.height
      });
    }, "image/jpeg", 0.96);
  });
}

function toPdfNumber(value) {
  return Number(value.toFixed(2)).toString();
}

function buildPdfBlob(pageImages) {
  const objectCount = 2 + pageImages.length * 3;
  const objects = new Array(objectCount + 1);
  const pageRefs = [];

  let objectIndex = 3;

  pageImages.forEach((pageImage, index) => {
    const pageObjectNumber = objectIndex;
    const contentObjectNumber = objectIndex + 1;
    const imageObjectNumber = objectIndex + 2;
    const imageName = `Im${index + 1}`;
    const contentStream = textEncoder.encode(
      `q\n${toPdfNumber(PDF_PAGE_WIDTH_PT)} 0 0 ${toPdfNumber(PDF_PAGE_HEIGHT_PT)} 0 0 cm\n/${imageName} Do\nQ`
    );

    pageRefs.push(`${pageObjectNumber} 0 R`);

    objects[pageObjectNumber] = textEncoder.encode(
      `${pageObjectNumber} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${toPdfNumber(PDF_PAGE_WIDTH_PT)} ${toPdfNumber(PDF_PAGE_HEIGHT_PT)}] /Resources << /XObject << /${imageName} ${imageObjectNumber} 0 R >> >> >> /Contents ${contentObjectNumber} 0 R >>\nendobj\n`
    );

    objects[contentObjectNumber] = concatUint8Arrays([
      textEncoder.encode(`${contentObjectNumber} 0 obj\n<< /Length ${contentStream.length} >>\nstream\n`),
      contentStream,
      textEncoder.encode(`\nendstream\nendobj\n`)
    ]);

    objects[imageObjectNumber] = concatUint8Arrays([
      textEncoder.encode(`${imageObjectNumber} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${pageImage.width} /Height ${pageImage.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${pageImage.bytes.length} >>\nstream\n`),
      pageImage.bytes,
      textEncoder.encode(`\nendstream\nendobj\n`)
    ]);

    objectIndex += 3;
  });

  objects[1] = textEncoder.encode(`1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n`);
  objects[2] = textEncoder.encode(`2 0 obj\n<< /Type /Pages /Count ${pageImages.length} /Kids [${pageRefs.join(" ")}] >>\nendobj\n`);

  const header = new Uint8Array([37, 80, 68, 70, 45, 49, 46, 52, 10, 37, 226, 227, 207, 211, 10]);
  const parts = [header];
  const offsets = new Array(objectCount + 1).fill(0);
  let offset = header.length;

  for (let index = 1; index <= objectCount; index += 1) {
    offsets[index] = offset;
    parts.push(objects[index]);
    offset += objects[index].length;
  }

  const xrefOffset = offset;
  const xrefLines = [
    `xref\n0 ${objectCount + 1}\n`,
    "0000000000 65535 f \n"
  ];

  for (let index = 1; index <= objectCount; index += 1) {
    xrefLines.push(`${String(offsets[index]).padStart(10, "0")} 00000 n \n`);
  }

  const trailer = `trailer\n<< /Size ${objectCount + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return new Blob([...parts, textEncoder.encode(`${xrefLines.join("")}${trailer}`)], {
    type: "application/pdf"
  });
}

async function exportPagesToPdf(pageElements, filename, label) {
  exportInProgress = true;
  refreshExportControls();

  try {
    const pageImages = [];

    for (let index = 0; index < pageElements.length; index += 1) {
      setExportStatus(`${label}：正在處理第 ${index + 1} / ${pageElements.length} 頁`, "");

      const canvas = await renderPageToCanvas(pageElements[index]);
      const pageImage = await canvasToJpegBytes(canvas);
      pageImages.push(pageImage);
      canvas.width = 0;
      canvas.height = 0;
    }

    const pdfBlob = buildPdfBlob(pageImages);
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
