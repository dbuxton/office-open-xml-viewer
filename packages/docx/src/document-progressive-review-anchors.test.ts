import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { DocxDocument } from './document.js';
import { resolvePageCommentAnchors } from './comment-margin.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { normalizeLayoutOptions } from './layout/options.js';
import { layoutDocumentProgressively } from './layout/progressive.js';
import {
  attachDocumentLayoutRuntime,
  documentLayoutRuntimeOf,
} from './layout/runtime-state.js';
import { setDocumentLayoutValidation } from './layout/validation-policy.js';
import type { DocumentLayout } from './layout/types.js';
import { retainRenderWorkerDocumentLayout } from './render-worker-layout.js';
import { resolveRevisionAnchorRuns } from './revisions.js';
import {
  installStubCanvas,
  syntheticDocxModel,
  type SyntheticDocumentShape,
} from './testing/synthetic-document.js';
import { textRunsForPage } from './text-run-projection.js';

// ─────────────────────────────────────────────────────────────────────────────
// Review anchors (§17.13.4 comments, §17.13.5 revisions) are projected against
// whichever layout the variant store currently holds. Under progressive layout
// that is first a PROVISIONAL PREFIX and only later the authoritative document,
// which broke the projection twice over:
//
//   1. every paragraph past the prefix's cut has no projected text, which the
//      fallback rule read as "this content has no final-state geometry" — so
//      each of those comments was anchored to the nearest run INSIDE the
//      prefix, i.e. beside page 1;
//   2. the projection was memoised with no tie to the layout it described, so
//      when the authoritative layout replaced the prefix the prefix-era
//      anchors survived — the comment then resolved BOTH on its real page (via
//      its authored source) and on page 0 (via the stale fallback), and the
//      duplicate never went away.
//
// Exercised through the same variant-store wiring `DocxDocument.load` uses
// (`retainRenderWorkerDocumentLayout` + `store.prime`), because `load()` itself
// needs a Worker and WASM.
// ─────────────────────────────────────────────────────────────────────────────

const CURRENT_DATE_MS = 1_700_000_000_000;
const PARAGRAPHS = 300;
/** Far enough past any first-preview cut that the prefix cannot reach it. */
const COMMENTED_PARAGRAPH = 280;

beforeAll(() => {
  installStubCanvas();
});

afterAll(() => {
  setDocumentLayoutValidation(true);
});

function progressiveDocument(
  shape: SyntheticDocumentShape,
  options: { commentedParagraphIndices?: readonly number[] } = {},
) {
  const model = syntheticDocxModel(shape, { paragraphs: PARAGRAPHS, ...options });
  const source = layoutSourceStore(model);
  const services = createLayoutServices(source);
  const retained = retainRenderWorkerDocumentLayout(source, services, CURRENT_DATE_MS);
  const doc = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(doc, {
    _mode: 'main',
    _document: model,
    _source: source,
    _meta: null,
    // What `load()` sets before the first publication can be observed.
    _layoutComplete: false,
  });
  attachDocumentLayoutRuntime(doc, CURRENT_DATE_MS);
  const runtime = documentLayoutRuntimeOf(doc);
  runtime.services = services;
  const layoutOptions = normalizeLayoutOptions(undefined, CURRENT_DATE_MS);
  runtime.activeLayoutOptions = layoutOptions;
  return { doc, source, services, store: retained.layoutVariants, layoutOptions };
}

/** Run the real progressive chain, returning the first published prefix and
 *  the authoritative layout — the two states the viewer paints from. */
async function previewAndFullLayout(
  source: ReturnType<typeof layoutSourceStore>,
  services: ReturnType<typeof createLayoutServices>,
  layoutOptions: ReturnType<typeof normalizeLayoutOptions>,
) {
  let preview: DocumentLayout | null = null;
  const full = await layoutDocumentProgressively(
    source.bodyLayoutInput,
    services,
    layoutOptions,
    {
      hasPaginationFields: source.hasPaginationFields,
      onPreview: (published) => {
        preview ??= published.layout as DocumentLayout;
      },
    },
  );
  if (!preview) throw new Error('the fixture produced no progressive preview');
  return { preview: preview as DocumentLayout, full };
}

/** First page whose projected runs come from the given body paragraph. */
function pageContainingParagraph(layout: DocumentLayout, paragraphIndex: number): number {
  for (let page = 0; page < layout.pages.length; page += 1) {
    const found = textRunsForPage(layout, page, { scale: 1 }).some(
      (run) => run.source?.story === 'body' && run.source.path[0] === paragraphIndex,
    );
    if (found) return page;
  }
  throw new Error(`paragraph ${paragraphIndex} is not projected on any page`);
}

describe('review anchors follow the layout a progressive load publishes', () => {
  it('never anchors a late comment inside the prefix, before or after handover', async () => {
    const { doc, source, services, store, layoutOptions } = progressiveDocument('plain', {
      commentedParagraphIndices: [COMMENTED_PARAGRAPH],
    });
    const { preview, full } = await previewAndFullLayout(source, services, layoutOptions);
    // The fixture only means anything if the prefix really stops short.
    expect(preview.pages.length).toBeLessThan(full.pages.length);

    store.prime(layoutOptions, preview);
    const provisional = doc.commentAnchorRanges();
    expect(provisional.map(({ commentId }) => commentId))
      .toEqual([String(COMMENTED_PARAGRAPH)]);
    // Symptom 1: the comment belongs to a page nobody has laid out yet, so the
    // prefix's own pages must show nothing for it.
    for (let page = 0; page < preview.pages.length; page += 1) {
      expect(resolvePageCommentAnchors(
        provisional,
        textRunsForPage(preview, page, { scale: 1 }),
      )).toEqual([]);
    }

    // The authoritative layout replaces the prefix, exactly as `load()` does.
    store.prime(layoutOptions, full, true);
    Object.assign(doc, { _layoutComplete: true });
    const final = doc.commentAnchorRanges();
    expect(final).not.toBe(provisional);

    // Symptom 2: the anchor must now resolve on its own page, and ONLY there.
    const page0Runs = textRunsForPage(full, 0, { scale: 1 });
    expect(resolvePageCommentAnchors(final, page0Runs)).toEqual([]);
    const home = pageContainingParagraph(full, COMMENTED_PARAGRAPH);
    expect(home).toBeGreaterThan(0);
    expect(resolvePageCommentAnchors(final, textRunsForPage(full, home, { scale: 1 }))
      .map(({ anchor }) => anchor.commentId)).toEqual([String(COMMENTED_PARAGRAPH)]);
  }, 300_000);

  it('re-projects revision anchors when the authoritative layout lands', async () => {
    const { doc, source, services, store, layoutOptions } = progressiveDocument('tracked');
    const { preview, full } = await previewAndFullLayout(source, services, layoutOptions);
    expect(preview.pages.length).toBeLessThan(full.pages.length);

    store.prime(layoutOptions, preview);
    const provisional = doc.revisionAnchorRanges();
    expect(provisional.length).toBeGreaterThan(0);
    // No deletion past the prefix cut may borrow geometry from inside it.
    const prefixRuns = preview.pages.flatMap(
      (_page, index) => textRunsForPage(preview, index, { scale: 1 }),
    );
    const projectedPaths = new Set(prefixRuns.map((run) => run.source?.path[0]));
    for (const anchor of provisional) {
      if (projectedPaths.has(anchor.source.path[0])) continue;
      expect(resolveRevisionAnchorRuns(anchor, prefixRuns)).toEqual([]);
    }

    store.prime(layoutOptions, full, true);
    Object.assign(doc, { _layoutComplete: true });
    expect(doc.revisionAnchorRanges()).not.toBe(provisional);
  }, 300_000);

  it('memoises the projection for as long as the layout is unchanged', async () => {
    // A cache that recomputed on every read would turn each mounted slot's
    // comment redraw into a document-wide story walk.
    const { doc, source, services, store, layoutOptions } = progressiveDocument('plain', {
      commentedParagraphIndices: [COMMENTED_PARAGRAPH],
    });
    const { full } = await previewAndFullLayout(source, services, layoutOptions);
    store.prime(layoutOptions, full, true);
    Object.assign(doc, { _layoutComplete: true });

    expect(doc.commentAnchorRanges()).toBe(doc.commentAnchorRanges());
    expect(doc.revisionAnchorRanges()).toBe(doc.revisionAnchorRanges());
  }, 300_000);

  it('returns one stable empty array for a document with no comments', async () => {
    // `DocxScrollViewer._hasDisplayableComments` invalidates its id set by
    // ARRAY IDENTITY, so a fresh `[]` per call would rebuild it on every
    // layout-math read.
    const { doc, source, services, store, layoutOptions } = progressiveDocument('plain');
    const { full } = await previewAndFullLayout(source, services, layoutOptions);
    store.prime(layoutOptions, full, true);
    Object.assign(doc, { _layoutComplete: true });

    expect(doc.commentAnchorRanges()).toEqual([]);
    expect(doc.commentAnchorRanges()).toBe(doc.commentAnchorRanges());
  }, 300_000);
});
