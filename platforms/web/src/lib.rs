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
#[cfg(feature = "engine-htmlbox")]
pub mod canvas_backend_htmlbox;
#[cfg(feature = "gui")]
pub mod canvas_backend_widgets;
pub mod console;
pub mod crypto;
pub mod dom_parser;
pub mod encoding;
pub mod engine;
/// Which engine is live, chosen at run time. See `engine_select`.
pub mod engine_select;
#[cfg(feature = "gui")]
pub mod engine_widgets;
/// The other engine behind the same trait. Additive to `engine_widgets`: both
/// are compiled in, and `install()` decides which one is live.
#[cfg(feature = "engine-htmlbox")]
pub mod engine_htmlbox;

/// Getting a frame out of whichever engine is live. A STOPGAP — see the
/// module docs for the shape that survives an out-of-process browser.
#[cfg(feature = "gui")]
pub mod present;

/// The browser NAMED by this build, for the paths that need a concrete type.
///
/// `platforms/web` is native code that owns the relationship with a browser; it
/// registers the `web:*` host functions and passes the calls on. Which browser
/// receives them is chosen at run time by `engine_select` — this alias is only
/// for the handful of calls that cannot go through the data enum, and it
/// resolves to htmlbox whenever htmlbox is compiled in.
///
/// Most operations go through `engine::apply` because they are DATA — a node
/// id, a name, a string, and which engine receives them is a RUNTIME choice
/// (`engine_select`). A few cannot: registering an event listener hands over a
/// CLOSURE, and no data enum can carry one. Those call the browser directly,
/// through this alias — and an alias is a type, so it is the one part of the
/// swap that stays build-time. See `with_browser` for what that costs.
#[cfg(feature = "engine-htmlbox")]
pub type Browser = rhtmledit::types::Document;
#[cfg(all(feature = "gui", not(feature = "engine-htmlbox")))]
pub type Browser = vybe_widgets::dom::Document;

/// Proof, at COMPILE TIME, that both browsers offer the same WHATWG surface.
///
/// Never called. It names the methods generically through [`Browser`], so if
/// either engine renames one, drops one, or changes a signature, this stops
/// compiling under that engine's feature — which is the only way "they are
/// interchangeable" can be a fact rather than a hope. A test cannot check it:
/// a test only ever runs against the engine that was built.
#[cfg(feature = "gui")]
#[allow(dead_code)]
fn _both_browsers_are_whatwg(browser: &mut Browser) {
    let node = browser.create_element("div");
    let text = browser.create_text_node("x");
    browser.append_child(node, text);
    browser.set_attribute(node, "id", "x");
    let _ = browser.get_attribute(node, "id");
    let _ = browser.get_attribute_names(node);
    browser.remove_attribute(node, "id");
    let _ = browser.node_type(node);
    let _ = browser.node_name(node);
    let _ = browser.node_value(node);
    let _ = browser.is_connected(node);
    let _ = browser.child_nodes(node);
    let _ = browser.parent_node(node);
    let _ = browser.text_content(node);
    browser.set_text_content(node, "x");
    let _ = browser.get_element_by_id("x");
    let _ = browser.query_selector("div");
    let _ = browser.query_selector_all("div");
    let _ = browser.get_elements_by_tag_name("div");
    let _ = browser.clone_node(node, true);
    let _ = browser.get_bounding_client_rect(node);
    let _ = browser.get_style_property(node, "top");
    let _ = browser.computed_style_property(node, "top");
    browser.set_style_property(node, "top", "1em");
    let _ = browser.checked(node);
    browser.set_checked(node, true);
    browser.focus(node);
    let _ = browser.value(node);
    browser.set_value(node, "v");
    let _ = browser.title();
    browser.set_title("t");
    let _ = browser.is_element(node);
    let _ = browser.has_attribute(node, "id");
    let _ = browser.namespace_uri(node);
    let _ = browser.local_name(node);
    browser.show_dialog(node, true);
    let _ = browser.dialog_open(node);
    browser.close_dialog(node);

    // ── Traversal, selectors and the ChildNode/ParentNode mixins ──
    let _ = browser.parent_element(node);
    let _ = browser.first_child(node);
    let _ = browser.last_child(node);
    let _ = browser.has_child_nodes(node);
    let _ = browser.contains(node, node);
    let _ = browser.children(node);
    let _ = browser.first_element_child(node);
    let _ = browser.last_element_child(node);
    let _ = browser.child_element_count(node);
    let _ = browser.next_element_sibling(node);
    let _ = browser.previous_element_sibling(node);
    let _ = browser.next_sibling(node);
    let _ = browser.previous_sibling(node);
    let _ = browser.matches(node, "div");
    let _ = browser.closest(node, "div");
    let _ = browser.get_elements_by_class_name("a b");
    let _ = browser.has_attributes(node);
    let _ = browser.toggle_attribute(node, "hidden");
    let _ = browser.class_name(node);
    browser.set_class_name(node, "c");
    let _ = browser.id(node);
    browser.set_id(node, "i");
    let _ = browser.document_element();
    let _ = browser.head();
    let _ = browser.body();
    let _ = browser.tag_name(node);
    let _ = browser.offset_width(node);
    let _ = browser.offset_height(node);
    browser.class_list_add(node, "c");
    let _ = browser.class_list_contains(node, "c");
    browser.class_list_toggle(node, "c");
    browser.class_list_remove(node, "c");
    let _ = browser.remove_style_property(node, "top");
    let _ = browser.inner_html(node);
    let _ = browser.item_text(node, 0);
    let _ = browser.selected_index(node);
    browser.set_selected_index(node, 0);
    browser.add_item(node, "x");
    browser.set_item_text(node, 0, "y");
    browser.remove_item(node, 0);
    browser.clear_items(node);
    let _ = browser.text_data(node);
    browser.set_text_data(node, "d");
    let _ = browser.is_text_node(node);
    let _ = browser.is_comment_node(node);
    let _ = browser.is_character_data(node);
    let _ = browser.kind();

    // ── DocumentFragment (DOM §4.2.1) ──
    let fragment = browser.create_document_fragment();
    let _ = browser.is_document_fragment(fragment);
    let inside = browser.create_element("p");
    browser.append_child(fragment, inside);
    // Appending the fragment moves its CHILDREN across — the fragment itself
    // never lands in the tree.
    browser.append_child(node, fragment);

    // ── Node comparison and normalisation (DOM §4.4) ──
    browser.normalize(node);
    let _ = browser.is_equal_node(node, node);
    let _ = browser.compare_document_position(node, node);

    // ── The rest of ParentNode / ChildNode (DOM §4.2.6) ──
    //
    // These take a SLICE because the IDL takes a variadic `(Node or DOMString)…`
    // and every one of them inserts the whole run at a single point. Calling
    // `insert_before` in a loop is not the same operation: it re-resolves the
    // reference child each time, which is how a caller ends up with its nodes
    // reversed.
    browser.append(node, &[text]);
    browser.prepend(node, &[text]);
    browser.before(text, &[]);
    browser.after(text, &[]);
    browser.replace_with(text, &[]);

    // ── Serialisation and adjacent insertion (DOM Parsing §3) ──
    let _ = browser.outer_html(node);
    let _ = browser.insert_adjacent_element(node, "beforeend", text);
    browser.insert_adjacent_text(node, "beforeend", "t");

    // ── HTMLElement (HTML §3.2.6, §7.6, CSSOM-View §5) ──
    let _ = browser.dataset(node);
    let _ = browser.dataset_get(node, "k");
    browser.set_dataset(node, "k", "v");
    browser.remove_dataset(node, "k");
    let _ = browser.inner_text(node);
    let _ = browser.tab_index(node);
    browser.set_tab_index(node, 0);
    browser.click(node);
    browser.scroll_into_view(node);
    let _ = browser.get_client_rects(node);
    let _ = browser.offset_top(node);
    let _ = browser.offset_left(node);
    let _ = browser.offset_parent(node);

    // ── The namespaced attribute accessors that were missing their pair ──
    let _ = browser.has_attribute_ns(node, "urn:x", "a");
    browser.remove_attribute_ns(node, "urn:x", "a");
}

/// Borrow a browser document as a CONCRETE type.
///
/// The direct-call path, for the few things `DomOp` cannot carry — registering
/// an event listener hands over a closure, and no data enum holds one.
///
/// **It serves one engine, and which one is fixed at build time.** That is not
/// an oversight: the closure takes `&mut Browser`, a concrete type, so this
/// cannot dispatch between two different engines at run time the way
/// `engine::apply` can. `Browser` names htmlbox whenever htmlbox is compiled
/// in, so on a build with both engines this reaches htmlbox — regardless of
/// which one `engine_select` made live.
///
/// So it answers `None` when the live engine is not the one it names, rather
/// than borrowing the wrong document. Silently operating on the other engine's
/// tree is exactly the failure the canvas seam had: every call succeeds and
/// nothing lands where the caller can see it.
///
/// Anything that can be expressed as data should go through `engine::apply`
/// instead, which follows the runtime choice.
#[cfg(feature = "gui")]
pub fn with_browser<T>(
    document: engine::DocumentId,
    f: impl FnOnce(&mut Browser) -> T,
) -> Option<T> {
    #[cfg(feature = "engine-htmlbox")]
    {
        if engine_select::live() != Some(engine_select::Engine::HtmlBox) {
            return None;
        }
        engine_htmlbox::with_document(document, f)
    }
    #[cfg(not(feature = "engine-htmlbox"))]
    {
        if engine_select::live() != Some(engine_select::Engine::Widgets) {
            return None;
        }
        engine_widgets::with_document(document, f)
    }
}
pub mod fetch;
pub mod html;
/// WHATWG File System Access — `showOpenFilePicker` and friends. Behind `gui`
/// because a picker is the user agent's own chrome, which only the toolkit has.
#[cfg(feature = "gui")]
pub mod file_system_access;
pub mod timers;
pub mod ui_events;
pub mod url;
pub mod window;

use vybe_runtime::VM;

pub fn register(vm: &mut VM) {
    // Install the engine the `web:*` surface talks to, and the canvas painter
    // that goes with it. WHICH one is a runtime choice — `VYBE_ENGINE`, or
    // `engine_select::choose` before this runs — because `set_engine` and
    // `set_backend` are runtime slots and always were.
    //
    // This used to be two pairs of `#[cfg]` blocks installing in a fixed order,
    // so the last one compiled in always won and swapping engines meant a
    // rebuild. The ordering was doing the work of a switch.
    engine_select::install();

    #[cfg(feature = "gui")]
    file_system_access::register(vm);
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
