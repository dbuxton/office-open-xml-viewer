import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import type { CommentAnchorRange } from './comments.js';
import {
  FakeDocxEngine,
  installDom,
  makeContainer,
  type FakeEl,
} from './scroll-viewer-test-dom.js';
import type { DocxTextRunInfo } from './types.js';

// ─────────────────────────────────────────────────────────────────────────────
// Progressive layout hands the viewer a document whose page count GROWS: it
// mounts the provisional opening pages, then relays out when the authoritative
// layout lands. The virtualization math already takes the heights array fresh
// on every pass, so what needs pinning is the viewer's side of that contract —
// the scroll extent tracks the new page count, the mounted window is unchanged
// for pages the user is already looking at, and scroll position survives.
// ─────────────────────────────────────────────────────────────────────────────

const PAGE = [{ widthPt: 612, heightPt: 792 }];

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function spacerOf(container: FakeEl): FakeEl {
  return container.children[0].children[0].children[0];
}

describe('DocxScrollViewer — growing page count', () => {
  it('extends the scroll region when layout completes', () => {
    installDom();
    const container = makeContainer(700, 500);
    // Two provisional pages, as a preview publishes.
    const engine = new FakeDocxEngine(2, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );
    const provisionalHeight = parseFloat(spacerOf(container).style.height);
    expect(viewer.pageCount).toBe(2);
    expect(provisionalHeight).toBeGreaterThan(0);

    // The authoritative layout arrives.
    engine.setPageCount(80);
    viewer.relayout();

    expect(viewer.pageCount).toBe(80);
    const finalHeight = parseFloat(spacerOf(container).style.height);
    expect(finalHeight).toBeGreaterThan(provisionalHeight);
    viewer.destroy();
  });

  it('keeps the pages already on screen mounted across the handover', () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(2, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );
    const mountedBefore = viewer.topVisiblePage;
    expect(mountedBefore).toBe(0);

    engine.setPageCount(80);
    viewer.relayout();

    // Growing the document must not scroll the user somewhere else.
    expect(viewer.topVisiblePage).toBe(mountedBefore);
    viewer.destroy();
  });

  it('repaints pages in place when the layout underneath them is replaced', () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(2, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );
    const paintedProvisionally = engine.renderCalls.length;
    expect(paintedProvisionally).toBeGreaterThan(0);

    // A plain relayout must NOT repaint: that guard is what keeps scrolling
    // cheap.
    viewer.relayout();
    expect(engine.renderCalls.length).toBe(paintedProvisionally);

    // Replacing the layout must, because a page's content can change without
    // its index changing (a footer's PAGE/NUMPAGES total, for one).
    (viewer as unknown as { _invalidateRenderedSlots(): void })._invalidateRenderedSlots();
    engine.setPageCount(80);
    viewer.relayout();
    expect(engine.renderCalls.length).toBeGreaterThan(paintedProvisionally);
    viewer.destroy();
  });

  it('reports a borrowed engine as fully laid out', async () => {
    // fromDocument borrows an already-loaded document, and an engine injected by
    // an integrator may predate these members entirely; neither may throw.
    installDom();
    const engine = new FakeDocxEngine(3, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
    );
    expect(viewer.layoutComplete).toBe(true);
    await expect(viewer.whenLayoutComplete()).resolves.toBeUndefined();
    viewer.destroy();
  });

  it('re-fires onVisiblePageChange when the total grows without the index moving', () => {
    installDom();
    const fires: Array<[number, number, boolean]> = [];
    const engine = new FakeDocxEngine(2, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
      { onVisiblePageChange: (top, total, complete) => { fires.push([top, total, complete]); } },
    );
    expect(fires).toEqual([[0, 2, true]]);

    // The user has not scrolled — topIndex is still 0 — but the document grew.
    // An index-only latch would strand the indicator on the preview count.
    engine.setPageCount(80);
    viewer.relayout();
    expect(fires).toEqual([[0, 2, true], [0, 80, true]]);
    viewer.destroy();
  });

  it('does not fire when neither the index nor the total changed', () => {
    installDom();
    const fires: Array<[number, number]> = [];
    const engine = new FakeDocxEngine(4, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
      { onVisiblePageChange: (top, total) => { fires.push([top, total]); } },
    );
    const initial = fires.length;
    viewer.relayout();
    viewer.relayout();
    expect(fires.length).toBe(initial);
    viewer.destroy();
  });

  it('re-fires when the document shrinks', () => {
    // A tracked-changes view can have FEWER pages than the final view, so the
    // count moves down as well as up.
    installDom();
    const fires: Array<[number, number]> = [];
    const engine = new FakeDocxEngine(40, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
      { onVisiblePageChange: (top, total) => { fires.push([top, total]); } },
    );
    fires.length = 0;
    engine.setPageCount(3);
    viewer.relayout();
    expect(fires).toEqual([[0, 3]]);
    viewer.destroy();
  });

  it('moves the document to the markup variant when tracked changes are toggled', () => {
    installDom();
    const engine = new FakeDocxEngine(20, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
    );
    expect(engine.layoutViews).toEqual([]);

    viewer.setShowTrackedChanges(true);
    expect(engine.layoutViews).toEqual([
      { showTrackedChanges: true, currentDate: undefined },
    ]);
    viewer.destroy();
  });

  it('recycles out-of-range slots when the markup variant is shorter', () => {
    // Hiding vs showing deletions changes the page count. Toggling to a SHORTER
    // variant while scrolled deep used to leave slots asking for pages that no
    // longer exist, which surfaced as a RangeError and a blank page.
    installDom();
    const engine = new FakeDocxEngine(60, PAGE);
    const container = makeContainer(700, 500);
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );
    viewer.scrollToPage(50);
    expect(viewer.topVisiblePage).toBeGreaterThan(0);

    // The toggle shortens the document under the reader.
    engine.setPageCount(3);
    engine.renderCalls.length = 0;
    viewer.setShowTrackedChanges(true);

    expect(viewer.pageCount).toBe(3);
    // Every page requested AFTER the toggle must exist in the shorter variant.
    expect(engine.renderCalls.length).toBeGreaterThan(0);
    for (const call of engine.renderCalls) {
      expect(call.page).toBeLessThan(3);
    }
    viewer.destroy();
  });

  it('does not mount the whole document just because it grew', () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(2, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { overscan: 1 },
    );
    engine.setPageCount(400);
    viewer.relayout();

    const scrollHost = container.children[0].children[0];
    const canvases = scrollHost.children.filter(
      (child: FakeEl) => child.children.some((nested: FakeEl) => nested.tag === 'canvas'),
    );
    // Virtualization still applies: a 400-page document mounts a viewport's
    // worth of slots, not 400.
    expect(canvases.length).toBeGreaterThan(0);
    expect(canvases.length).toBeLessThan(10);
    viewer.destroy();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// A publication renumbers the document's pages, so everything the viewer keyed
// on a page index — the comment→page map, the retained per-page run geometry,
// and any navigation still scanning them — belongs to the layout that is being
// replaced. Left in place, `goToComment` kept steering to a page the comment
// had only ever occupied in the provisional prefix.
// ─────────────────────────────────────────────────────────────────────────────

const COMMENT_SOURCE = { story: 'body', storyInstance: 'body', path: [7] } as const;

function anchoredRun(): DocxTextRunInfo {
  return {
    text: 'annotated', source: COMMENT_SOURCE, sourceRunIndex: 0,
    x: 20, y: 200, w: 80, h: 14, fontSize: 12, font: '12px sans-serif',
  };
}

function commentedEngine(pages: number): FakeDocxEngine {
  const engine = new FakeDocxEngine(pages, PAGE);
  engine.comments = [{ id: 'c', author: 'Ada', text: 'Anchored late' }];
  engine.commentAnchors = [{
    commentId: 'c',
    source: COMMENT_SOURCE,
    startRunIndex: 0,
    endRunIndex: 1,
    reference: { source: COMMENT_SOURCE, runIndex: 1, affinity: 'preceding' },
  }] as CommentAnchorRange[];
  return engine;
}

/** The viewer's own response to a progressive publication, as `load()` wires
 *  it — unreachable through `fromDocument`, which borrows a loaded document. */
function publish(viewer: ReturnType<typeof DocxScrollViewer.fromDocument>): void {
  (viewer as unknown as { _onLayoutPublication(exact: boolean): void })
    ._onLayoutPublication(false);
}

describe('DocxScrollViewer — comment navigation across a publication', () => {
  it('re-scans comment pages once the authoritative layout lands', async () => {
    installDom();
    const engine = commentedEngine(3);
    // In the prefix the anchored run is projected on page 1.
    engine.collectPageRuns = vi.fn(async (page: number) => (
      page === 1 ? [anchoredRun()] : []));
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer() as unknown as HTMLElement,
      engine.asDoc(),
    );

    await expect(viewer.goToComment('c')).resolves.toBe(true);
    expect(viewer.getSelectionContext()).toMatchObject({ commentId: 'c', pageIndex: 1 });

    // The authoritative layout puts the same paragraph much further down.
    engine.setPageCount(60);
    engine.collectPageRuns = vi.fn(async (page: number) => (
      page === 41 ? [anchoredRun()] : []));
    publish(viewer);

    await expect(viewer.goToComment('c')).resolves.toBe(true);
    expect(viewer.getSelectionContext()).toMatchObject({ commentId: 'c', pageIndex: 41 });
    viewer.destroy();
  });

  it('abandons a navigation the publication renumbered under it', async () => {
    installDom();
    const engine = commentedEngine(3);
    let resolveFirstPage!: (runs: DocxTextRunInfo[]) => void;
    const firstPage = new Promise<DocxTextRunInfo[]>((resolve) => {
      resolveFirstPage = resolve;
    });
    engine.collectPageRuns = vi.fn((page: number) => (
      page === 0 ? firstPage : Promise.resolve([anchoredRun()])));
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer() as unknown as HTMLElement,
      engine.asDoc(),
    );

    const inFlight = viewer.goToComment('c');
    await Promise.resolve();
    publish(viewer);
    resolveFirstPage([]);

    // Its scan described pages the new layout has renumbered; the caller retries.
    await expect(inFlight).resolves.toBe(false);
    viewer.destroy();
  });
});

describe('DocxScrollViewer — a borrowed document that is still laying out', () => {
  it('re-scans comment pages when the authoritative layout lands', async () => {
    // `load({ progressiveLayout })` resolves on the opening pages, so an
    // application that owns its DocxDocument — the composition `fromDocument`
    // exists for — hands the viewer a PREFIX. Nothing else tells the viewer the
    // authoritative layout replaced it, so the prefix's page count and its
    // prefix-era comment index stayed on screen for good.
    installDom();
    const engine = commentedEngine(2);
    engine.beginProgressiveLayout();
    engine.collectPageRuns = vi.fn(async (page: number) => (
      page === 1 ? [anchoredRun()] : []));
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer() as unknown as HTMLElement,
      engine.asDoc(),
    );

    expect(viewer.layoutComplete).toBe(false);
    await expect(viewer.goToComment('c')).resolves.toBe(true);
    expect(viewer.getSelectionContext()).toMatchObject({ commentId: 'c', pageIndex: 1 });

    engine.collectPageRuns = vi.fn(async (page: number) => (
      page === 41 ? [anchoredRun()] : []));
    engine.completeLayout(60);
    await engine.whenLayoutComplete();
    await Promise.resolve();

    expect(viewer.pageCount).toBe(60);
    await expect(viewer.goToComment('c')).resolves.toBe(true);
    expect(viewer.getSelectionContext()).toMatchObject({ commentId: 'c', pageIndex: 41 });
    viewer.destroy();
  });
});

describe('DocxScrollViewer — comment chrome stays legible when zoomed out', () => {
  function zoomOf(viewer: ReturnType<typeof DocxScrollViewer.fromDocument>): number {
    return (viewer as unknown as { _commentZoom(): number })._commentZoom();
  }

  it('floors the comment zoom so cards do not shrink with the page', () => {
    installDom();
    const engine = commentedEngine(4);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
      { comments: true },
    );

    viewer.setScale(0.35);
    // The page is at 0.35; the cards are chrome, not content, and stay readable.
    expect(zoomOf(viewer)).toBe(1);
    // Zooming IN still scales them with the document.
    viewer.setScale(2);
    expect(zoomOf(viewer)).toBe(2);
    viewer.destroy();
  });

  it('lets an application tie the chrome back to the document zoom', () => {
    installDom();
    const engine = commentedEngine(4);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
      { comments: { minZoom: 0 } },
    );

    viewer.setScale(0.35);
    expect(zoomOf(viewer)).toBeCloseTo(0.35, 5);
    viewer.destroy();
  });
});
