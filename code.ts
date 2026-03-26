figma.showUI(__html__, { width: 280, height: 240 });

const VN_STOCKS = [
  'VCB', 'HPG', 'VIC', 'VNM', 'TCB', 'FPT', 'MSN', 'VHM', 'GAS', 'MWG',
  'VPB', 'MBB', 'ACB', 'VJC', 'POW', 'SSI', 'HDB', 'PLX', 'VRE', 'STB',
  'GVR', 'BID', 'CTG', 'VIB', 'TPB', 'SHB', 'LPB', 'EIB', 'MSB', 'DGC'
];

type SmartColumnType = 'stock' | 'date' | 'volume' | 'report';

/**
 * Hàm tìm kiếm đệ quy tất cả các TextNode bên trong một Node
 */
function findAllTextNodes(node: SceneNode): TextNode[] {
  let textNodes: TextNode[] = [];
  if (node.type === 'TEXT') {
    textNodes.push(node);
  } else if ('children' in node) {
    for (const child of node.children) {
      textNodes = textNodes.concat(findAllTextNodes(child as SceneNode));
    }
  }
  return textNodes;
}

function getAllTextContent(node: SceneNode): string {
  const textNodes = findAllTextNodes(node);
  return textNodes.map((t) => t.characters).join(' ').trim();
}

function detectColumnType(headerTextRaw: string): SmartColumnType | null {
  const headerText = headerTextRaw.trim().toLowerCase();
  if (headerText.includes('mã ck')) return 'stock';
  if (headerText.includes('ngày tạo')) return 'date';
  if (headerText.includes('khối lượng')) return 'volume';
  if (headerText.includes('tên báo cáo')) return 'report';
  return null;
}

/**
 * Hàm lấy text hiển thị từ node đầu tiên (thường là Header)
 */
function getHeaderText(node: SceneNode): string {
  return getAllTextContent(node);
}

async function loadFontsForTextNode(textNode: TextNode): Promise<void> {
  if (textNode.fontName !== figma.mixed) {
    await figma.loadFontAsync(textNode.fontName as FontName);
    return;
  }

  const fonts: FontName[] = [];
  const len = textNode.characters.length;
  for (let i = 0; i < len; i++) {
    const font = textNode.getRangeFontName(i, i + 1);
    if (font === figma.mixed) continue;
    const f = font as FontName;
    const exists = fonts.some((x) => x.family === f.family && x.style === f.style);
    if (!exists) fonts.push(f);
  }
  for (const font of fonts) {
    await figma.loadFontAsync(font);
  }
}

function formatNumberWithCommas(value: number): string {
  const s = Math.floor(value).toString();
  return s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function isPlaceholderText(text: string): boolean {
  const t = text.trim();
  if (t === '') return true;
  const normalized = t.toLowerCase();
  return normalized === '-' || normalized === '—' || normalized === 'n/a' || normalized === 'na';
}

function isHeaderLabelText(text: string): boolean {
  const t = text.trim().toLowerCase();
  return t === 'mã ck' || t === 'ngày tạo' || t === 'khối lượng' || t === 'tên báo cáo';
}

function getFontSizeScore(textNode: TextNode): number {
  return typeof textNode.fontSize === 'number' ? textNode.fontSize : 0;
}

function pickCellValueTextNode(textNodes: TextNode[], dataType: SmartColumnType): TextNode | null {
  if (textNodes.length === 0) return null;

  const candidates = textNodes.filter((t) => !isHeaderLabelText(t.characters));
  const list = candidates.length > 0 ? candidates : textNodes;

  const matchers: Record<SmartColumnType, (s: string) => boolean> = {
    stock: (s) => isPlaceholderText(s) || /^[A-Z]{2,6}$/.test(s.trim()),
    date: (s) => isPlaceholderText(s) || /^\d{2}\/\d{2}\/\d{4}$/.test(s.trim()),
    volume: (s) => isPlaceholderText(s) || /^[\d,]+$/.test(s.trim()),
    report: (s) => isPlaceholderText(s) || s.trim().toUpperCase().startsWith('THEO DÕI')
  };

  const direct = list.find((t) => matchers[dataType](t.characters));
  if (direct) return direct;

  return list.reduce((best, current) => {
    const bestSize = getFontSizeScore(best);
    const currentSize = getFontSizeScore(current);
    if (currentSize !== bestSize) return currentSize > bestSize ? current : best;
    return current.characters.length > best.characters.length ? current : best;
  }, list[0]);
}

/**
 * Generator dữ liệu ngẫu nhiên
 */
const RandomData = {
  stock: () => VN_STOCKS[Math.floor(Math.random() * VN_STOCKS.length)],
  date: () => {
    const start = new Date(2023, 0, 1);
    const end = new Date();
    const date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
    return `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}`;
  },
  volume: () => {
    const val = Math.floor(Math.random() * 2000) * 500; // Bước nhảy 500
    return formatNumberWithCommas(val);
  },
  reportName: (stockOverride?: string) => {
    const stock =
      stockOverride !== undefined && stockOverride !== null && stockOverride !== ''
        ? stockOverride
        : VN_STOCKS[Math.floor(Math.random() * VN_STOCKS.length)];
    const date = new Date();
    const dateStr = `${date.getDate().toString().padStart(2, '0')}${ (date.getMonth() + 1).toString().padStart(2, '0')}${date.getFullYear().toString().slice(-2)}`;
    return `THEO DÕI - Báo cáo định giá - ${stock} - ${dateStr}`;
  }
};

function isPotentialColumn(node: SceneNode): node is FrameNode {
  if (node.type !== 'FRAME') return false;
  if (!('children' in node) || node.children.length <= 1) return false;
  if (node.layoutMode !== 'VERTICAL') return false;
  const headerCell = node.children[0] as SceneNode;
  const headerText = getHeaderText(headerCell);
  return detectColumnType(headerText) !== null;
}

function collectColumnsFromSelection(selection: readonly SceneNode[]): FrameNode[] {
  const seen = new Set<string>();
  const columns: FrameNode[] = [];

  const addColumn = (col: FrameNode) => {
    if (seen.has(col.id)) return;
    seen.add(col.id);
    columns.push(col);
  };

  for (const node of selection) {
    if (isPotentialColumn(node)) addColumn(node);
    if ('findAll' in node) {
      const nodeWithFindAll = node as unknown as {
        findAll: (callback: (n: SceneNode) => boolean) => SceneNode[];
      };
      const found = nodeWithFindAll.findAll((n: SceneNode) => isPotentialColumn(n));
      for (const col of found) addColumn(col as FrameNode);
    }
  }

  return columns;
}

figma.ui.onmessage = async (msg) => {
  if (msg.type === 'fill-smart-data') {
    const selection = figma.currentPage.selection;

    if (selection.length === 0) {
      figma.notify('Vui lòng chọn ít nhất một cột (Auto Layout Frame)');
      return;
    }

    const columns = collectColumnsFromSelection(selection);
    if (columns.length === 0) {
      figma.notify('Không tìm thấy cột phù hợp (cần chọn cột Auto Layout theo chiều dọc).');
      return;
    }

    let updatedCount = 0;

    const columnInfos = columns
      .map((col) => {
        const headerCell = col.children[0] as SceneNode;
        const headerTitle = getHeaderText(headerCell);
        const dataType = detectColumnType(headerTitle);
        return { col, dataType };
      })
      .filter((x): x is { col: FrameNode, dataType: SmartColumnType } => x.dataType !== null);

    const stockByRowIndex = new Map<number, string>();

    for (const { col, dataType } of columnInfos.filter((x) => x.dataType === 'stock')) {
      const children = col.children;
      for (let i = 1; i < children.length; i++) {
        const cell = children[i] as SceneNode;
        const cellValueText = pickCellValueTextNode(findAllTextNodes(cell), dataType);
        if (!cellValueText) continue;
        await loadFontsForTextNode(cellValueText);

        const stock = RandomData.stock();
        stockByRowIndex.set(i, stock);
        cellValueText.characters = stock;
        updatedCount++;
      }
    }

    for (const { col, dataType } of columnInfos.filter((x) => x.dataType !== 'stock')) {
      const children = col.children;
      for (let i = 1; i < children.length; i++) {
        const cell = children[i] as SceneNode;
        const cellValueText = pickCellValueTextNode(findAllTextNodes(cell), dataType);
        if (!cellValueText) continue;
        await loadFontsForTextNode(cellValueText);

        if (dataType === 'date') cellValueText.characters = RandomData.date();
        else if (dataType === 'volume') cellValueText.characters = RandomData.volume();
        else if (dataType === 'report') {
          const sameRowStock = stockByRowIndex.get(i);
          cellValueText.characters = RandomData.reportName(sameRowStock);
        }

        updatedCount++;
      }
    }

    if (updatedCount > 0) {
      figma.notify(`Đã điền thông minh dữ liệu cho ${updatedCount} ô.`);
    } else {
      figma.notify('Không tìm thấy cột phù hợp để điền dữ liệu.');
    }
  }
};
