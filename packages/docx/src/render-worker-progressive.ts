/**
 * Progressive pagination inside the render worker.
 *
 * ## Why this is a separate module
 *
 * Worker mode was built as a single RPC: one `parse`, a blocking
 * `doc.layoutVariants.defaultLayout`, then one `parsedMeta` carrying the whole
 * document's geometry. That is a fine answer to "don't freeze the UI thread",
 * but it is the wrong answer to "show me something now" — the host waits for
 * the LAST page before it can paint the first.
 *
 * Everything progressive layout needs was already worker-safe. Fonts are
 * preloaded into `self.fonts` before pagination measures anything, and
 * `drainPaginationAsync` yields through a `MessageChannel`, which exists in
 * workers too. What was missing was somewhere to put a provisional layout and a
 * message to announce it. This module supplies the first half; the wire's
 * `layoutPartial` supplies the second.
 *
 * It is a separate module rather than more code in `render-worker.ts` for two
 * reasons. It has no `self` or WASM dependency, so tests exercise it directly
 * (see `render-worker-progressive.test.ts`) instead of standing up a Worker.
 * And `render-worker.ts`'s metadata route is AST-pinned by
 * `scripts/check-docx-layout-boundaries.mjs` to exactly one `layout` /
 * `pageSizes` / `meta` declaration, so the partials' geometry has to be built
 * somewhere else.
 *
 * ## Why the worker's metadata route stays a single line
 *
 * This primes the AUTHORITATIVE layout into the variant store before it
 * returns, under the very options key the caller passed. The worker's
 * `const layout = doc.layoutVariants.layoutFor(layoutOptions)` therefore
 * becomes a cache hit rather than a second pagination: identical bytes, no
 * repeated work, and one spelling of "the layout this parse is reporting on".
 *
 * `layoutOptions` is supplied rather than derived here precisely so it cannot
 * drift from the one the metadata route selects. Priming under a key nothing
 * reads would be worse than not previewing at all: the whole progressive pass
 * would be discarded AND a second full layout built.
 */
import { buildBookmarkPageMap } from './bookmark-nav.js';
import { layoutDocumentProgressively } from './layout/progressive.js';
import type { DeepReadonly, DocumentLayout } from './layout/types.js';
import type { LayoutOptions } from './layout/options.js';
import type { LayoutSourceStore } from './layout/layout-source-store.js';
import type { RetainedRenderWorkerDocumentLayout } from './render-worker-layout.js';
import type { DocumentLayoutPartial } from './worker-protocol.js';

/** One provisional publication: the geometry a host needs to grow its page
 *  count, its `pageSize` answers and its scroll extent. Structurally the wire's
 *  {@link DocumentLayoutPartial} minus the review payload, which the caller
 *  attaches to the first publication only. */
export type RenderWorkerLayoutPublication = Omit<DocumentLayoutPartial, 'review'>;

/** How a publication and its progress reach the outside world. Injected rather
 *  than imported so this module never touches `self` or the wire, which is what
 *  lets `render-worker-progressive.test.ts` drive it under plain vitest. */
export interface RenderWorkerLayoutPublisher {
  /** A provisional prefix has been primed and is safe to request pages from. */
  publish(publication: RenderWorkerLayoutPublication): void;
  /** Committed pages so far, at every pagination suspension point. Fires far
   *  too often to forward verbatim — the caller throttles. */
  progress(committedPages: number): void;
}

/** Project a published prefix into the geometry the host consumes. Mirrors the
 *  worker's authoritative metadata route (page sizes from stamped canonical
 *  geometry, bookmarks from the same paginated pages) so a partial and the
 *  final `parsedMeta` describe pages the same way. */
function publicationOf(
  layout: DocumentLayout | DeepReadonly<DocumentLayout>,
  exact: boolean,
): RenderWorkerLayoutPublication {
  return {
    pageCount: layout.pages.length,
    pageSizes: layout.pages.map((page) => ({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    })),
    bookmarkPages: [...buildBookmarkPageMap(layout)],
    exact,
  };
}

/**
 * Lay the document out progressively, priming every step into the retained
 * variant store and announcing the provisional ones through `publish`.
 *
 * Resolves once the authoritative layout is primed. Only PREVIEWS are
 * published: the authoritative layout reaches the host as the `parse`
 * response's `parsedMeta`, so publishing it here too would just describe the
 * same pages twice.
 *
 * ## Prime before publish
 *
 * A publication tells the host that more pages exist, and the host will
 * immediately ask to render them. Priming first is therefore not an ordering
 * preference but a correctness requirement: it guarantees the store can serve
 * every page the host has been invited to request, so `requireLayoutPage`
 * cannot raise a `RangeError` for a page the host was told about.
 */
export async function paginateRenderWorkerDocumentProgressively(
  doc: RetainedRenderWorkerDocumentLayout,
  source: LayoutSourceStore,
  publisher: RenderWorkerLayoutPublisher,
  layoutOptions: LayoutOptions,
  signal?: AbortSignal,
): Promise<void> {
  const store = doc.layoutVariants;
  let published = false;
  const layout = await layoutDocumentProgressively(
    source.bodyLayoutInput,
    doc.layoutServices,
    layoutOptions,
    {
      hasPaginationFields: source.hasPaginationFields,
      scheduler: { signal, onProgress: (committedPages) => publisher.progress(committedPages) },
      onPreview: (preview) => {
        // `replace` is false for the first publication (nothing to replace) and
        // true afterwards, matching the main-thread path: a provisional prefix
        // is the one layout a later pass is allowed to supersede.
        store.prime(layoutOptions, preview.layout, published);
        published = true;
        publisher.publish(publicationOf(preview.layout, preview.exact));
      },
    },
  );
  // Replaces the provisional prefix. Anything the host already painted from it
  // is stale by design — it repaints when the authoritative metadata lands.
  store.prime(layoutOptions, layout, true);
}
