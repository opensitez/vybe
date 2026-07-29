//! # JSPI (JS Promise Integration) custom section
//!
//! Proposal: <https://github.com/WebAssembly/js-promise-integration>.
//! JSPI has **no new WASM opcodes** — it's entirely a JS-API-layer
//! feature. A conforming JS host wraps specific exports /imports with
//! `WebAssembly.promising(fn)` / `new WebAssembly.Suspending(fn)`; at
//! that point async-looking JS ↔ sync-looking wasm works transparently.
//!
//! The problem: the WASM binary itself has no standard way to say "this
//! export is promising." The spec defers that to out-of-band metadata.
//! Vybe emits that metadata as a custom section named `vybe.jspi` so
//! the JS glue loader can pick it up without guesswork.
//!
//! # Layout of the `vybe.jspi` custom section
//!
//! ```text
//! payload:
//!     u32 promising_count           — number of promising exports
//!     u32 promising_func_idx * N    — WASM function indices to wrap
//!     u32 suspending_count          — number of suspending imports
//!     u32 suspending_import_idx * M — WASM import indices to wrap
//! ```
//!
//! Function indices follow WASM's convention: imports come first
//! (0..import_count), user-defined chunks follow (import_count..).
//!
//! `compiler_common` sets `Chunk::is_async = true` on every `async`
//! function the user writes; that chunk's function index becomes a
//! `promising_func_idx`. Import wrapping is identified by Vybe's
//! convention: any host fn named `wasm:js-*.await*` is listed as a
//! `suspending_import_idx`.
//!
//! # Wire format choice: custom section vs. export name
//!
//! The simpler alternative ("mark promising exports by appending
//! `[promising]` to the export name") would tunnel the signal through
//! ExportDesc and survive any WASM engine. It's uglier for DevTools
//! and pollutes the export namespace visible to plain JS callers, so
//! we chose the custom section. Engines that don't understand it
//! silently ignore it (per WASM spec, custom sections MUST be
//! ignored when the consumer doesn't know the name).

use crate::encoding::*;
use vybe_runtime::Chunk;

pub const SECTION_NAME: &str = "vybe.jspi";

/// Build the payload for the `vybe.jspi` custom section. Returns
/// `None` when no chunk is marked async — callers can then skip
/// emitting the section entirely.
pub fn encode_payload(chunks: &[Chunk], rt_imports_len: usize) -> Option<Vec<u8>> {
    let host_imports_len = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
    let import_base = host_imports_len + rt_imports_len;

    let promising: Vec<u32> = chunks
        .iter()
        .enumerate()
        // Async GENERATORS are continuations, not JSPI-promising calls —
        // `is_async` is source truth, so exclude generators here.
        .filter(|(_, c)| c.is_async && !c.is_generator)
        .map(|(i, _)| (import_base + i) as u32)
        .collect();

    let suspending: Vec<u32> = chunks
        .first()
        .map(|c| {
            c.imports
                .iter()
                .enumerate()
                .filter(|(_, import)| is_suspending_import(&import.module, &import.name))
                .map(|(i, _)| i as u32)
                .collect()
        })
        .unwrap_or_default();

    if promising.is_empty() && suspending.is_empty() {
        return None;
    }

    let mut out = Vec::new();
    write_leb128_u32(&mut out, promising.len() as u32);
    for idx in &promising {
        write_leb128_u32(&mut out, *idx);
    }
    write_leb128_u32(&mut out, suspending.len() as u32);
    for idx in &suspending {
        write_leb128_u32(&mut out, *idx);
    }
    Some(out)
}

fn is_suspending_import(module: &str, name: &str) -> bool {
    module.starts_with("wasm:js-") && name.contains("await")
}
