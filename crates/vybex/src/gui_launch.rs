//! GUI launch layer — owns the winit event loop, VM glue, and dialog registration.
//!
//! Uses `widgets::Form` as the container for all controls, and
//! `widgets::Application` + `run_app()` for the window/event loop.
//! All graphics, focus management, hover states, and keyboard routing live
//! in widgets.
//!
//! One entry point: `launch_gui`.
//!
//! There were two. `launch_vybewidget_form` built a form by CONSTRUCTING
//! widgets from a WinForms designer model — `Button`, `DataGrid`, `TreeView`
//! imported by Rust type — which is `document.createElement` spelled as a type
//! import, in the crate that is supposed to be the user agent rather than the
//! page. It sat behind a `gui_forms` feature that no manifest in the workspace
//! ever enabled, so it compiled out of every build; 344 lines of it, deleted
//! rather than left to look like a supported path.

use std::cell::RefCell;
use std::rc::Rc;

use widgets::{
    Application, FontSystem, KeyEvent, MouseEvent, PanelWidget, Pixmap, SwashCache, run_app,
};


// ── FormApp — Application impl ─────────────────────────────────────────

struct FormApp {
    font_system: FontSystem,
    swash_cache: SwashCache,
    vm: Rc<RefCell<vybe_runtime::VM>>,
    initialised: bool,
}

impl Application for FormApp {
    fn on_init(&mut self, width: f32, height: f32, _scale: f32) {
        Self::lay_out(width, height);
        if !self.initialised {
            self.initialised = true;
            self.fire_load_event();
        }
    }

    fn on_resize(&mut self, width: f32, height: f32) {
        Self::lay_out(width, height);
    }

    /// **The window's title is the document's title.**
    ///
    /// `document.title = "Calculator"` reached `Document::set_title` and stopped
    /// there: `FormApp` never implemented this, so it returned the trait's empty
    /// default, `app_window` skips an empty title, and the window kept the name
    /// the launcher gave it at startup — `Form1`, whatever the page said.
    ///
    /// Read live rather than cached at launch, which is what makes a title set
    /// LATER — or changed while running — reach the chrome. `app_window` polls
    /// this after each render and only touches the window when the string
    /// actually changes, so re-reading per frame costs a comparison.
    fn title(&self) -> String {
        crate::gui_document::with_live(|doc| doc.title()).unwrap_or_default()
    }

    fn render(&mut self, pixmap: &mut Pixmap, scale: f32) {
        // One renderer, shared with `--capture` and the debugger's `capture`,
        // so a captured frame is byte-for-byte what the window shows.
        crate::gui_capture::render_into(pixmap, scale);
    }

    fn handle_mouse(&mut self, event: MouseEvent) -> bool {
        if Self::gui_trace_enabled() {
            eprintln!("[gui] formapp.handle_mouse event={:?}", event);
        }
        // Window events become W3C UI Events in the `web:ui-events` queue —
        // the same queue a browser host fills from the real DOM. SDL's
        // vocabulary is applied later, by SDL's own adapter, not here.
        {
            use widgets::layout::{MouseButton, MouseEventKind};
            use widgets::ui_events::{UiEvent, queue};
            // DOM `button`: 0 left, 1 middle, 2 right.
            let dom_button = |b: &MouseButton| match b {
                MouseButton::Left => 0i32,
                MouseButton::Middle => 1,
                MouseButton::Right => 2,
            };
            // DOM `buttons` mask: 1 left, 2 right, 4 middle.
            let dom_mask = |b: &MouseButton| match b {
                MouseButton::Left => 1i32,
                MouseButton::Right => 2,
                MouseButton::Middle => 4,
            };
            let (kind, button, buttons) = match &event.kind {
                MouseEventKind::Press(b) => ("mousedown", dom_button(b), dom_mask(b)),
                MouseEventKind::Release(b) => ("mouseup", dom_button(b), 0),
                MouseEventKind::Move | MouseEventKind::Scroll(_) => ("mousemove", 0, 0),
            };
            queue().push(UiEvent {
                kind: kind.to_string(),
                client_x: event.x as i32,
                client_y: event.y as i32,
                button,
                buttons,
                ..UiEvent::default()
            });
            // Through the seam, so the LIVE engine hit-tests it. This called
            // `widgets`' form directly, which is the toolkit whether or
            // not the toolkit is the engine in use — so under
            // `--engine webcore` every click was hit-tested against an empty
            // tree and nothing at all happened.
            //
            // The SAME W3C fields the queue above is given: one translation
            // from the window's vocabulary, two consumers.
            vybe_platform_web::engine::apply(
                crate::gui_document::active(),
                vybe_platform_web::engine::DomOp::DispatchPointer {
                    kind: kind.to_string(),
                    client_x: event.x,
                    client_y: event.y,
                    button,
                },
            );
        }
        self.dispatch_document_events();
        true
    }

    fn handle_key(&mut self, event: KeyEvent) -> bool {
        // Both edges as `keydown`/`keyup`, in W3C shape.
        {
            use widgets::ui_events::{UiEvent, queue};
            let pressed = event.state == widgets::winit::event::ElementState::Pressed;
            let (key, code, key_code) = dom_key_fields(&event.key_without_modifiers);
            queue().push(UiEvent {
                kind: if pressed { "keydown" } else { "keyup" }.to_string(),
                key,
                code,
                key_code,
                ctrl_key: event.cmd,
                shift_key: event.shift,
                alt_key: event.alt,
                ..UiEvent::default()
            });
            // Widget dispatch is pressed-only — releases exist for the event
            // queue alone, so typing/focus behavior is unchanged.
            if !pressed {
                return false;
            }
        }
        crate::gui_document::with_live(|d| d.form_mut().handle_key(&event));
        self.dispatch_document_events();
        true
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        {
            use widgets::ui_events::{UiEvent, queue};
            // DOM `deltaY` is positive DOWN — the opposite of a scroll delta
            // that reports "up" as positive.
            queue().push(UiEvent {
                kind: "wheel".to_string(),
                client_x: x as i32,
                client_y: y as i32,
                delta_y: -(delta as f64),
                ..UiEvent::default()
            });
        }
        let handled = crate::gui_document::with_live(|d| d.form_mut().handle_scroll(delta, x, y))
            .unwrap_or(false);
        self.dispatch_document_events();
        handled
    }

    fn cursor_icon(&self) -> widgets::CursorIcon {
        widgets::CursorIcon::Default
    }

    /// The frame callback, called ~60 Hz from the event loop's
    /// `about_to_wait`.
    fn on_tick(&mut self) {
        // ── The frame boundary ────────────────────────────────────────────
        //
        // `requestAnimationFrame` callbacks run HERE, before the ~60 Hz
        // repaint, which is precisely the browser's contract: a page is
        // called back before the next paint and draws then. This is what
        // replaces "present the buffer" — the window is already redrawing,
        // so a guest that re-registers each frame gets a real animation
        // loop, and one that stops asking costs nothing.
        {
            use vybe_runtime::scheduler::DeferredSource;
            let clock = vybe_platform_web::animation::callbacks();
            let stamp = vybe_runtime::Value::F64(vybe_runtime::event_loop::monotonic_now_ms());
            // Drain only what this frame owes: `pop_due` stops handing out
            // callbacks once the frame advances, so a callback that
            // re-registers runs NEXT frame instead of spinning here.
            while let Some(cb) = clock.pop_due() {
                let mut vm = self.vm.borrow_mut();
                let _ = match fn_arity(&cb) {
                    0 => vm.invoke(&cb, &[]),
                    _ => vm.invoke(&cb, &[stamp.clone()]),
                };
            }
        }

        // Anything the document queued that input handling did not already
        // drain — a widget that reports on a later frame, or an event raised
        // by a timer/animation callback rather than by a click.
        self.dispatch_document_events();
    }
}

// ── VM glue ────────────────────────────────────────────────────────────

impl FormApp {
    fn gui_trace_enabled() -> bool {
        std::env::var("VYBE_GUI_TRACE")
            .map(|value| !matches!(value.as_str(), "" | "0" | "false" | "False"))
            .unwrap_or(false)
    }

    /// Give the tree the window's size. The document's viewport IS the body's
    /// containing block, so a document that never gets one lays every control
    /// out against its 800×600 default instead of the window.
    fn lay_out(width: f32, height: f32) {
        crate::gui_document::with_live(|d| d.set_viewport(width, height));
    }

    /// Turn what the user did in the document into VM calls.
    ///
    /// `web:dom` listeners are where EVERY frontend's event wiring ends up —
    /// `primitives/gui.rs` lowers VCL's `OnClick := h`, WinForms' `Click += h`
    /// and Flutter's `onPressed: h` to the same `addEventListener` — so this
    /// single drain serves all of them and none of it is framework-specific.
    ///
    /// A listener is called with the Event and nothing else, which is what a
    /// document does and the only thing it CAN do — the receiver is already
    /// bound into the handler by `primitives/gui.rs`, so there is no form
    /// object to look up and no arity table to keep in step with four
    /// frontends.
    fn dispatch_document_events(&mut self) {
        for dispatch in crate::gui_document::drain() {
            if Self::gui_trace_enabled() {
                eprintln!(
                    "[gui] document dispatch kind={} sender={}",
                    dispatch.kind, dispatch.sender
                );
            }
            let mut vm = self.vm.borrow_mut();
            if let Err(e) = vm.invoke(&dispatch.callback, &[dispatch.event]) {
                eprintln!("Event handler error: {e}");
            }
        }
    }

    /// `Handles Me.Load` is a subscription on the FORM, and a form IS the
    /// document's body — so the listener lives on the body node and this is a
    /// plain `load` dispatch, exactly what a page does.
    fn fire_load_event(&mut self) {
        let event = crate::gui_document::event_object("load", widgets::dom::DOCUMENT);
        for listener in crate::gui_document::listeners_for(widgets::dom::DOCUMENT, "load") {
            let mut vm = self.vm.borrow_mut();
            if let Err(e) = vm.invoke(&listener, &[event.clone()]) {
                eprintln!("[LOAD] Error: {e}");
            }
        }
    }

}

// ── Helpers ────────────────────────────────────────────────────────────

pub(crate) fn fn_arity(val: &vybe_runtime::Value) -> usize {
    match val {
        vybe_runtime::Value::Object(obj) => match &obj.lock().unwrap().kind {
            vybe_runtime::value::ObjectKind::Function(f) => f.arity as usize,
            _ => 0,
        },
        _ => 0,
    }
}

// ── Dialog registration ────────────────────────────────────────────────

/// winit key → W3C `KeyboardEvent` fields: `(key, code, keyCode)`.
///
/// `key` is what the keypress MEANS ("a", "Enter", "ArrowLeft"), `code` is
/// the physical key ("KeyA", "Digit1", "ArrowLeft") and stays layout-
/// independent, `keyCode` is the legacy numeric identity browsers still
/// ship. No SDL here — SDL's keysyms are derived from these by its own
/// adapter, so a browser host producing real DOM events needs no changes.
fn dom_key_fields(key: &widgets::winit::keyboard::Key) -> (String, String, i32) {
    use widgets::winit::keyboard::{Key, NamedKey};
    match key {
        Key::Character(text) => {
            let Some(c) = text.chars().next() else {
                return (String::new(), String::new(), 0);
            };
            let lower = c.to_ascii_lowercase();
            let code = match lower {
                'a'..='z' => format!("Key{}", lower.to_ascii_uppercase()),
                '0'..='9' => format!("Digit{}", lower),
                '-' => "Minus".to_string(),
                '=' => "Equal".to_string(),
                ' ' => "Space".to_string(),
                _ => String::new(),
            };
            // `keyCode` is the uppercase code point for letters, the digit
            // for digits — the browser's legacy convention.
            let key_code = if lower.is_ascii_alphabetic() {
                lower.to_ascii_uppercase() as i32
            } else {
                lower as i32
            };
            (c.to_string(), code, key_code)
        }
        Key::Named(named) => {
            let f = |k: &str, c: &str, kc: i32| (k.to_string(), c.to_string(), kc);
            match named {
                NamedKey::Enter => f("Enter", "Enter", 13),
                NamedKey::Escape => f("Escape", "Escape", 27),
                NamedKey::Backspace => f("Backspace", "Backspace", 8),
                NamedKey::Tab => f("Tab", "Tab", 9),
                NamedKey::Space => f(" ", "Space", 32),
                NamedKey::Delete => f("Delete", "Delete", 46),
                NamedKey::ArrowRight => f("ArrowRight", "ArrowRight", 39),
                NamedKey::ArrowLeft => f("ArrowLeft", "ArrowLeft", 37),
                NamedKey::ArrowDown => f("ArrowDown", "ArrowDown", 40),
                NamedKey::ArrowUp => f("ArrowUp", "ArrowUp", 38),
                NamedKey::Home => f("Home", "Home", 36),
                NamedKey::End => f("End", "End", 35),
                NamedKey::PageUp => f("PageUp", "PageUp", 33),
                NamedKey::PageDown => f("PageDown", "PageDown", 34),
                NamedKey::Insert => f("Insert", "Insert", 45),
                NamedKey::CapsLock => f("CapsLock", "CapsLock", 20),
                NamedKey::F1 => f("F1", "F1", 112),
                NamedKey::F2 => f("F2", "F2", 113),
                NamedKey::F3 => f("F3", "F3", 114),
                NamedKey::F4 => f("F4", "F4", 115),
                NamedKey::F5 => f("F5", "F5", 116),
                NamedKey::F6 => f("F6", "F6", 117),
                NamedKey::F7 => f("F7", "F7", 118),
                NamedKey::F8 => f("F8", "F8", 119),
                NamedKey::F9 => f("F9", "F9", 120),
                NamedKey::F10 => f("F10", "F10", 121),
                NamedKey::F11 => f("F11", "F11", 122),
                NamedKey::F12 => f("F12", "F12", 123),
                NamedKey::Control => f("Control", "ControlLeft", 17),
                NamedKey::Shift => f("Shift", "ShiftLeft", 16),
                NamedKey::Alt => f("Alt", "AltLeft", 18),
                _ => (String::new(), String::new(), 0),
            }
        }
        _ => (String::new(), String::new(), 0),
    }
}

// ── Public launch functions ────────────────────────────────────────────

/// The window's size is the document's viewport — a form's `Width`/`Height`
/// are CSS on the body and land there. A document that never got one keeps the
/// 800×600 initial containing block.
fn window_size() -> (u32, u32) {
    crate::gui_document::viewport().unwrap_or((800, 600))
}

/// Open a window on the live document and run the event loop.
pub fn launch_gui(vm: vybe_runtime::VM) {
    let (width, height) = window_size();

    let app = FormApp {
        font_system: FontSystem::new(),
        swash_cache: SwashCache::new(),
        vm: Rc::new(RefCell::new(vm)),
        initialised: false,
    };

    run_app("Form1", width, height, 1.0, app);
}

/// Headless counterpart to `launch_gui`: lay the document out, render ONE frame
/// offscreen, write it as a PNG, and return — no window, no event loop.
///
/// Used by `--capture`; it makes a GUI program's output readable without a
/// screenshot, which also means a frame can be diffed as a regression check.
pub fn capture_gui(
    vm: vybe_runtime::VM,
    path: &str,
    control: Option<&str>,
) -> Result<(u32, u32), String> {
    let (width, height) = window_size();

    let mut app = FormApp {
        font_system: FontSystem::new(),
        swash_cache: SwashCache::new(),
        vm: Rc::new(RefCell::new(vm)),
        initialised: false,
    };

    // Lays the document out and fires `load`, exactly as the window path does —
    // without this, controls have no rect and the frame is blank.
    app.on_init(width as f32, height as f32, 1.0);

    crate::gui_capture::capture_to_png(path, control, 1.0)
}



