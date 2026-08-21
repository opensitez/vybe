//! WinForms application and window verbs, lowered to the web platform.
//!
//! `Application.Run` and the layout no-ops emit nothing at all;
//! `MessageBox.Show` is `window.alert`; `Form.Close` / `Application.Exit` /
//! `Form.Activate` / `Form.CenterToScreen` reach `web:window` through
//! `activeDocument().defaultView`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype};

/// HTML's `Window` — `alert`/`confirm` and the window verbs, all registered in
/// `platforms/web/src/window.rs`.
const WINDOW_MODULE: &str = "web:window";

/// `MessageBox.Show(text[, caption[, buttons…]])` → `window.alert(text)`.
///
/// **A browser's `alert()` takes a MESSAGE and nothing else.** HTML gives the
/// dialog's title to the user agent — a page cannot set it, deliberately, so
/// that a page cannot dress a prompt up as a system dialog. WinForms' `caption`
/// is therefore not "not implemented yet", it is UNREPRESENTABLE, and the
/// honest lowering drops it.
///
/// ⛔ It is NOT folded into the message. `alert(caption & vbCrLf & text)` would
/// put the title back on screen and would be a shim — the exact move that makes
/// a swapped-in engine render something no browser would.
///
/// Nothing is lost against what actually ran: the retired implementation read
/// `text` and `title`, showed an info box and returned null, so the BUTTONS
/// argument was already ignored and the result was already not a
/// `DialogResult`.
///
/// ⚠ `MessageBoxButtons.OKCancel` / `.YesNo` want `window.confirm`, which is
/// registered on `web:window` and returns a real `Bool`. That is a genuine
/// improvement and a genuine BEHAVIOUR CHANGE — the return goes from null to a
/// boolean, which a `DialogResult` comparison would see — so it is left for its
/// own change rather than smuggled in here.
///
/// `text` is pushed FIRST and so sits DEEPEST; the extras above it are dropped.
pub fn emit_message_box_show(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    // `Show()` with no arguments is not a WinForms overload, but an emit that
    // assumes its argument is present would underflow the stack rather than say
    // so. Give `alert` the empty string it would otherwise be missing.
    if argc == 0 {
        chunk.emit_string_const("", line);
    }
    let idx = chunk.add_import(WINDOW_MODULE, "alert");
    chunks[current].emit_call(idx, 1, line);
}

/// `Application.Run(form)` — **the event loop is the user agent's**.
///
/// A page is not told to run. It runs because it HAS a document: the browsing
/// context is the window, the UA owns the loop, and a script that finishes
/// leaves a live page behind. There is no `Application.Run` in the web platform
/// to be compliant WITH, so the compliant lowering is to emit no call at all —
/// the same answer `dotnet.self` gives, for the same reason.
///
/// The three things the retired implementation did are all answered by the
/// document now:
///
/// - "should this present?" → `should_present` asks `has_browsing_context()`
///   FIRST. A WinForms program reaches `activeDocument` to build any control at
///   all, so it has a context and still presents.
/// - The window's size, read off the form object's `width`/`height` → those are
///   CSS on the body, and `gui_launch` already prefers
///   `gui_document::viewport()` with `GuiState`'s pair only as the fallback
///   "for a form built without a document". Verified, not assumed: TicTacToe
///   captures at its declared 220x320 `ClientSize`, NOT the 800x600
///   `GuiState::new` default — so the viewport was already supplying it.
/// - `form_object` (seeds the `__f` global, the receiver a HOST-invoked handler
///   gets) → dead for a converted frontend. Events reach `web:dom`'s
///   `addEventListener`, and that path invokes the callback with a DOM event
///   object without consulting `__f` at all.
///
/// ⚠ The arguments are still EVALUATED — `Application.Run(New Form1())` builds
/// its form, and that construction is the whole program. Only the call goes
/// away: the args are dropped and one value is pushed, because every host call
/// pushes exactly one value and the statement that wraps this emits one `DROP`.
pub fn emit_application_run(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_ref_null(heaptype::HT_EXTERN, line);
}

/// `Form.Close()` → `window.close()` on this document's browsing context.
///
/// A form IS a window (`Form` maps to `body`, and the body's document sits in a
/// browsing context), so closing one is HTML §7.2.2's `close()` and not a node
/// operation. The window is named the standard way — `activeDocument()` then
/// `defaultView` — because a guest has no global object to spell `window` with.
///
/// ⚠ This does not yet make the app EXIT, and neither did what it replaces,
/// which set `GuiState::close_requested` — a field WRITTEN in three places and
/// READ IN NONE, so `Form.Close()` has long been a
/// no-op. The converted form is strictly more than that: it marks the context
/// closed in a registry `window.closed` can actually read. Making the shell act
/// on it belongs to `gui_launch`, which owns the event loop.
///
/// ⚠ In a real engine `close()` only acts on a SCRIPT-CLOSABLE window — one
/// opened by script, or a top-level traversable with a single history entry —
/// and is otherwise ignored. So a faithful implementation closing the main
/// window is a judgement call about which of those our top-level context is,
/// not an obviously-correct default.
pub fn emit_form_close(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_window_verb(chunks, current, argc, "close", line);
}

/// `Form.Activate()` → `window.focus()`.
///
/// Activating a form is bringing its WINDOW forward, which is HTML's
/// `window.focus()`. Deliberately not the control-level
/// `gui.ctrl.bring_to_front`: that re-appends a NODE to change document order,
/// and a form is not a node among siblings — it is the context they all live
/// in. Two different operations on two different things that a shared name
/// would have quietly merged.
pub fn emit_form_activate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_window_verb(chunks, current, argc, "focus", line);
}

/// `Form.CenterToScreen()` → `window.moveTo((screen - inner) / 2, …)`.
///
/// The arithmetic is the whole verb, and it belongs in bytecode: a browser has
/// `moveTo` and `screen`, and "centred" is what you get by subtracting the
/// window from the display and halving. No host function is needed to divide
/// by two.
///
/// ⛔ The retired implementation hardcoded **1920x1080** — its own comment
/// called that "a sensible default… common enough… for most
/// users" — and then wrote `Left`/`Top` into the `GuiState` property store,
/// which a converted frontend does not read. It also read `GuiState`'s width
/// and height, which are the 800x600 constructor defaults now that
/// `runApplication` no longer sets them. So it computed a guess from a stale
/// number and stored it where nothing looked.
///
/// Unset, `screenWidth`/`screenHeight` answer the VIEWPORT, so this resolves to
/// `moveTo(0, 0)` — "already centred", which is the truthful answer for a
/// display we cannot measure. `vybe_widgets::window::set_screen` is the door for
/// a shell that knows the real monitor.
pub fn emit_form_center_to_screen(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let win = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    emit_default_view(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, win, line);

    // `moveTo(win, x, y)` — the handle first, then the two coordinates.
    chunk.emit_op_u16(Op::LOCAL_GET, win, line);
    for (screen, inner) in [
        ("screenWidth", "innerWidth"),
        ("screenHeight", "innerHeight"),
    ] {
        for name in [screen, inner] {
            chunk.emit_op_u16(Op::LOCAL_GET, win, line);
            let idx = chunk.add_import(WINDOW_MODULE, name);
            chunk.emit_call(idx, 1, line);
        }
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_f64_const(2.0, line);
        chunk.emit_op(Op::F64_DIV, line);
    }
    let move_to = chunk.add_import(WINDOW_MODULE, "moveTo");
    chunk.emit_call(move_to, 3, line);
}

/// `Form.ShowDialog()` → `DialogResult.OK`, and nothing else.
///
/// **Modality is not implemented, and was not before this.** The retired
/// implementation did exactly two things: set `GuiState::should_run`, and
/// return `I32(1)`. Its own comment said so —
/// *"real modal handling is a separate workstream (it requires nested message
/// loops)"*. So this is not a capability being dropped; it is the same
/// non-capability, off the host.
///
/// `should_run` needs no replacement for the same reason `Application.Run`
/// needed none: the document already has a browsing context, and
/// `should_present` asks that first. The return value is the whole observable
/// behaviour, so the return value is the whole emit.
///
/// ⚠ **Where this actually goes.** A modal is `<dialog>.showModal()`, which IS
/// registered on `web:html` — but a `Form` maps to `body`, and `body` is not a
/// `<dialog>`. Wiring it needs forms to be real windows first (a second form is
/// a second browsing context), which is the multi-window model. Emitting a
/// constant is the honest placeholder; calling `showModal` on the body would be
/// a shim that looks like progress.
pub fn emit_form_show_dialog(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    // `DialogResult.OK`. An `i32`, matching the `Value::I32(1)` the host fn
    // returned — a float here would compare differently against the enum.
    chunk.emit_i32_const(1, line);
}

/// `activeDocument().defaultView` — this context's window, left on the stack.
fn emit_default_view(chunk: &mut Chunk, line: u32) {
    let doc = chunk.add_import(
        vybe_compiler::primitives::gui::DOCUMENT_MODULE,
        vybe_compiler::primitives::gui::HOST_FN_ACTIVE_DOCUMENT,
    );
    chunk.emit_call(doc, 0, line);
    let default_view = chunk.add_import(
        vybe_compiler::primitives::gui::DOCUMENT_MODULE,
        "defaultView",
    );
    chunk.emit_call(default_view, 1, line);
}

/// `<receiver>.<verb>()` → `activeDocument().defaultView.<verb>()`.
///
/// The receiver is dropped: a form does not name its own window — the DOCUMENT
/// does, and `defaultView` is the standard accessor for it. A guest cannot
/// spell `window`/`self`/`globalThis` because it has no global object, so this
/// two-hop is how any window verb is reached.
fn emit_window_verb(chunks: &mut [Chunk], current: usize, argc: u8, verb: &str, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    emit_default_view(chunk, line);
    let verb_idx = chunk.add_import(WINDOW_MODULE, verb);
    chunk.emit_call(verb_idx, 1, line);
}

/// `Application.Exit()` → `window.close()` on this browsing context.
///
/// Ending a WinForms app is ending its message loop, and the loop belongs to
/// the browsing context — so exiting IS closing the context, the same operation
/// `Form.Close` performs, reached the same way. There is no separate "quit" in
/// the web platform because there is no separate application: the page is the
/// program.
///
/// Same caveat as [`emit_form_close`]: the context is marked closed, and
/// nothing yet acts on that. What it replaces set `GuiState::close_requested`,
/// which has three writers and no readers, so this is not a behaviour it used
/// to have and lost.
pub fn emit_application_exit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_form_close(chunks, current, argc, line);
}

/// `SuspendLayout` / `ResumeLayout` / `PerformLayout` — declared, and doing
/// nothing.
///
/// These are real WinForms methods that must RESOLVE (an undeclared method is
/// `Me.SuspendLayout()` reaching `undefined`, which kills every designer's
/// `InitializeComponent` on its first line) but have nothing to do: layout is
/// the document's job and it is never batched by the guest. The host fn they
/// called was `Box::new(|_ctx, _| Value::Null)` — a registered name whose
/// entire body was "return null".
///
/// A const push is exactly equivalent under the one-value-per-host-call rule,
/// so this needs no host at all.
pub fn emit_noop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_ref_null(heaptype::HT_EXTERN, line);
}

