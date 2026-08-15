//! # `web:*` host imports — WHATWG / W3C web platform APIs.
//!
//! Web platform APIs that complement ECMA-262 but live outside the
//! language spec. Every JavaScript runtime ships these; Vybe exposes
//! them under the `web:*` namespace so language profiles can target a
//! standard surface instead of inventing `vybe:*` adapters.
//!
//! - `web:crypto`         — Web Cryptography API (`crypto.randomUUID`,
//!                          `crypto.getRandomValues`, `crypto.subtle.digest`)
//! - `web:url`            — WHATWG URL (`URL`, `URLSearchParams`)
//! - `web:encoding`       — WHATWG Encoding (`TextEncoder`, `TextDecoder`)
//! - `web:fetch`          — WHATWG Fetch (`fetch`, `Request`, `Response`,
//!                          `Headers`)
//! - `web:timers`         — HTML Timers (`setTimeout`, `clearTimeout`,
//!                          `setInterval`, `clearInterval`)
//! - `web:html`           — WHATWG DOM + HTML: `createElement` (unparented
//!                          until `appendChild`), attributes (`id` is an
//!                          attribute, matched by `getElementById`), CSS
//!                          `style.setProperty` with units, IDL `value` /
//!                          `checked`, and `addEventListener` whose
//!                          listeners receive an `Event` object. THE TREE
//!                          lives here; a renderer follows it via
//!                          `TreeObserver` (MutationObserver-shaped).
//! - `web:animation`      — HTML `requestAnimationFrame` /
//!                          `cancelAnimationFrame` + `performance.now`.
//!                          THE FRAME CLOCK — what a page uses instead of
//!                          presenting a buffer.
//! - `web:canvas`         — WHATWG HTML `CanvasRenderingContext2D`
//!                          (`getContext`, `fillRect`, `fillText`, `arc`,
//!                          `setLineDash`, `drawImage`). Paints through a
//!                          swappable backend — see `canvas_backend`.
//! - `web:ui-events`      — W3C UI Events (`KeyboardEvent`, `MouseEvent`,
//!                          `WheelEvent`): `dispatchEvent`, `pollEvent`,
//!                          `pointerState`. THE EVENT QUEUE LIVES HERE —
//!                          a native window backend pushes into it, a
//!                          browser host would fill it from the real DOM.
//! - `web:dom-parser`     — WHATWG DOM Parsing and Serialization
//!                          (`DOMParser.parseFromString`,
//!                          `XMLSerializer.serializeToString`) —
//!                          currently exposed as the flat 3-fn surface
//!                          `parse(s)` / `load(url)` / `toString(node)`
//!                          pending full Document/Element resource types.

pub mod animation;
pub mod builtin_types; // TypeRegistry vtables for the web surface; run in Plugin::finalize
pub mod canvas;
pub mod canvas_backend;
#[cfg(feature = "gui")]
pub mod canvas_backend_widgets;
pub mod console;
pub mod crypto;
pub mod dom_parser;
pub mod encoding;
pub mod engine;
#[cfg(feature = "gui")]
pub mod engine_widgets;
pub mod fetch;
pub mod html;
pub mod timers;
pub mod ui_events;
pub mod url;
pub mod window;

use vybe_runtime::VM;

pub fn register(vm: &mut VM) {
    // Install the engine the `web:*` surface talks to — windows, documents,
    // input. `vybe_widgets` is that engine today; a build without the `gui`
    // feature simply has none, the way a canvas with no painter draws nothing.
    #[cfg(feature = "gui")]
    engine_widgets::install();
    // The painter for `web:canvas`, resolving through that same document.
    // Installed HERE rather than by the vybe platform, which used to resolve
    // through `GuiState` — a second widget tree that a `createElement` canvas
    // is not in.
    #[cfg(feature = "gui")]
    canvas_backend_widgets::install();

    console::register(vm);
    crypto::register(vm);
    url::register(vm);
    encoding::register(vm);
    fetch::register(vm);
    timers::register(vm);
    dom_parser::register(vm);
    ui_events::register(vm);
    canvas::register(vm);
    animation::register(vm);
    html::register(vm);
    window::register(vm);
}
pub mod plugin;
pub use plugin::Plugin;
