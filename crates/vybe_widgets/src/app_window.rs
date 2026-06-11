//! AppWindow — window management and event loop for GUI applications.
//!
//! Handles all winit + softbuffer boilerplate. Apps implement the
//! `Application` trait and call `run_app()`.
//!
//! ```ignore
//! struct MyApp { /* state */ }
//! impl Application for MyApp { /* ... */ }
//! run_app("My App", 1200, 900, 2.0, MyApp::new());
//! ```

use std::num::NonZeroU32;
use std::sync::Arc;

use softbuffer::{Context, Surface};
use tiny_skia::Pixmap;
use winit::application::ApplicationHandler;
use winit::event::{ElementState, WindowEvent};
use winit::event_loop::{ControlFlow, EventLoop};
use winit::window::{CursorIcon, Window, WindowAttributes};

use crate::layout::{KeyEvent, MouseButton, MouseEvent, MouseEventKind};

fn gui_trace_enabled() -> bool {
    std::env::var("VYBE_GUI_TRACE")
        .map(|value| !matches!(value.as_str(), "" | "0" | "false" | "False"))
        .unwrap_or(false)
}

/// Trait for GUI applications built on the vybe_widgets toolkit.
///
/// The toolkit handles all window management, event translation, and surface
/// blitting. The application just renders into a pixmap and handles events.
pub trait Application {
    /// Called once after the window is created. `width`/`height` are logical pixels.
    fn on_init(&mut self, width: f32, height: f32, scale: f32);

    /// Called when the window is resized (logical dimensions).
    fn on_resize(&mut self, width: f32, height: f32);

    /// Render the entire UI into the pixmap.
    fn render(&mut self, pixmap: &mut Pixmap, scale: f32);

    /// Handle a mouse event. Return `true` to request a redraw.
    fn handle_mouse(&mut self, event: MouseEvent) -> bool;

    /// Handle a keyboard event. Return `true` to request a redraw.
    fn handle_key(&mut self, event: KeyEvent) -> bool;

    /// Handle scroll input. `delta` is positive for scroll-up. Return `true` to redraw.
    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool;

    /// Return the desired cursor icon for the current state.
    fn cursor_icon(&self) -> CursorIcon {
        CursorIcon::Default
    }

    /// Called when the window loses focus or cursor leaves.
    fn on_focus_lost(&mut self) {}

    /// Return a window title to display. Called after each render; the window
    /// title is updated only when the returned string changes.
    fn title(&self) -> String {
        String::new()
    }
}

// ── Internal Runner ────────────────────────────────────────────────────

struct AppWindowInner<A: Application> {
    app: A,
    window: Option<Arc<Window>>,
    context: Option<Context<Arc<Window>>>,
    surface: Option<Surface<Arc<Window>, Arc<Window>>>,
    pixmap: Option<Pixmap>,
    scale: f32,
    title: String,
    init_width: u32,
    init_height: u32,
    modifiers: winit::event::Modifiers,
    mouse_pos: (f32, f32),
}

impl<A: Application> ApplicationHandler for AppWindowInner<A> {
    fn resumed(&mut self, event_loop: &winit::event_loop::ActiveEventLoop) {
        let window = Arc::new(
            event_loop
                .create_window(
                    WindowAttributes::default()
                        .with_title(&self.title)
                        .with_inner_size(winit::dpi::LogicalSize::new(
                            self.init_width as f64,
                            self.init_height as f64,
                        )),
                )
                .unwrap(),
        );
        let ctx = Context::new(window.clone()).unwrap();
        let surf = Surface::new(&ctx, window.clone()).unwrap();
        let sz = window.inner_size();

        // Use the actual display scale factor from the OS
        self.scale = window.scale_factor() as f32;

        self.window = Some(window);
        self.context = Some(ctx);
        self.surface = Some(surf);
        self.pixmap = Some(Pixmap::new(sz.width.max(1), sz.height.max(1)).unwrap());

        let s = self.surface.as_mut().unwrap();
        s.resize(
            NonZeroU32::new(sz.width.max(1)).unwrap(),
            NonZeroU32::new(sz.height.max(1)).unwrap(),
        )
        .unwrap();

        let lw = sz.width as f32 / self.scale;
        let lh = sz.height as f32 / self.scale;
        self.app.on_init(lw, lh, self.scale);

        self.window.as_ref().unwrap().request_redraw();
    }

    fn window_event(
        &mut self,
        event_loop: &winit::event_loop::ActiveEventLoop,
        _id: winit::window::WindowId,
        event: WindowEvent,
    ) {
        let scale = self.scale;
        let mods = self.modifiers.state();

        match event {
            WindowEvent::CloseRequested => event_loop.exit(),

            WindowEvent::ScaleFactorChanged { scale_factor, .. } => {
                self.scale = scale_factor as f32;
                if let Some(w) = &self.window {
                    let sz = w.inner_size();
                    if sz.width > 0 && sz.height > 0 {
                        if let Some(s) = &mut self.surface {
                            s.resize(
                                NonZeroU32::new(sz.width).unwrap(),
                                NonZeroU32::new(sz.height).unwrap(),
                            )
                            .unwrap();
                        }
                        self.pixmap = Some(Pixmap::new(sz.width, sz.height).unwrap());
                        self.app
                            .on_resize(sz.width as f32 / self.scale, sz.height as f32 / self.scale);
                        w.request_redraw();
                    }
                }
            }

            WindowEvent::ModifiersChanged(m) => {
                self.modifiers = m;
            }

            WindowEvent::Resized(sz) => {
                if sz.width > 0 && sz.height > 0 {
                    if let Some(s) = &mut self.surface {
                        s.resize(
                            NonZeroU32::new(sz.width).unwrap(),
                            NonZeroU32::new(sz.height).unwrap(),
                        )
                        .unwrap();
                    }
                    self.pixmap = Some(Pixmap::new(sz.width, sz.height).unwrap());
                    self.app
                        .on_resize(sz.width as f32 / scale, sz.height as f32 / scale);
                    if let Some(w) = &self.window {
                        w.request_redraw();
                    }
                }
            }

            WindowEvent::MouseWheel { delta, .. } => {
                let d = match delta {
                    winit::event::MouseScrollDelta::LineDelta(_, y) => y * 120.0,
                    winit::event::MouseScrollDelta::PixelDelta(pos) => pos.y as f32 * 2.0,
                };
                let x = self.mouse_pos.0 / scale;
                let y = self.mouse_pos.1 / scale;
                if self.app.handle_scroll(d, x, y) {
                    if let Some(w) = &self.window {
                        w.request_redraw();
                    }
                }
            }

            WindowEvent::KeyboardInput { event, .. } => {
                if event.state == ElementState::Pressed {
                    #[cfg(target_os = "macos")]
                    let key_without_mods = {
                        use winit::platform::modifier_supplement::KeyEventExtModifierSupplement;
                        event.key_without_modifiers()
                    };
                    #[cfg(not(target_os = "macos"))]
                    let key_without_mods = event.logical_key.clone();

                    let key_event = KeyEvent {
                        logical_key: event.logical_key.clone(),
                        key_without_modifiers: key_without_mods,
                        state: event.state,
                        cmd: mods.super_key() || mods.control_key(),
                        shift: mods.shift_key(),
                        alt: mods.alt_key(),
                        text: event.text.as_ref().map(|t| t.to_string()),
                    };
                    if self.app.handle_key(key_event) {
                        if let Some(w) = &self.window {
                            w.request_redraw();
                        }
                    }
                }
            }

            WindowEvent::CursorMoved { position, .. } => {
                self.mouse_pos = (position.x as f32, position.y as f32);
                let event = MouseEvent {
                    x: position.x as f32 / scale,
                    y: position.y as f32 / scale,
                    kind: MouseEventKind::Move,
                    cmd: mods.super_key() || mods.control_key(),
                    shift: mods.shift_key(),
                    alt: mods.alt_key(),
                };
                if gui_trace_enabled() {
                    eprintln!(
                        "[gui] window.cursor_moved physical=({:.1},{:.1}) logical=({:.1},{:.1}) scale={:.2}",
                        position.x, position.y, event.x, event.y, scale,
                    );
                }
                if self.app.handle_mouse(event) {
                    if let Some(w) = &self.window {
                        w.request_redraw();
                    }
                }
                // Update cursor
                let icon = self.app.cursor_icon();
                if let Some(w) = &self.window {
                    w.set_cursor(icon);
                }
            }

            WindowEvent::Focused(false) | WindowEvent::CursorLeft { .. } => {
                self.app.on_focus_lost();
                if let Some(w) = &self.window {
                    w.request_redraw();
                }
            }

            WindowEvent::MouseInput { state, button, .. } => {
                let btn = match button {
                    winit::event::MouseButton::Left => MouseButton::Left,
                    winit::event::MouseButton::Right => MouseButton::Right,
                    winit::event::MouseButton::Middle => MouseButton::Middle,
                    _ => return,
                };
                let kind = match state {
                    ElementState::Pressed => MouseEventKind::Press(btn),
                    ElementState::Released => MouseEventKind::Release(btn),
                };
                let event = MouseEvent {
                    x: self.mouse_pos.0 / scale,
                    y: self.mouse_pos.1 / scale,
                    kind,
                    cmd: mods.super_key() || mods.control_key(),
                    shift: mods.shift_key(),
                    alt: mods.alt_key(),
                };
                if gui_trace_enabled() {
                    eprintln!(
                        "[gui] window.mouse_input state={:?} button={:?} logical=({:.1},{:.1}) physical=({:.1},{:.1}) scale={:.2}",
                        state, button, event.x, event.y, self.mouse_pos.0, self.mouse_pos.1, scale,
                    );
                }
                if self.app.handle_mouse(event) {
                    if let Some(w) = &self.window {
                        w.request_redraw();
                    }
                }
            }

            WindowEvent::RedrawRequested => {
                if let (Some(pix), Some(surf)) = (self.pixmap.as_mut(), self.surface.as_mut()) {
                    self.app.render(pix, scale);
                    // Update window title when the app requests a change
                    let new_title = self.app.title();
                    if !new_title.is_empty() && new_title != self.title {
                        self.title = new_title.clone();
                        if let Some(w) = &self.window {
                            w.set_title(&new_title);
                        }
                    }
                    let mut buffer = surf.buffer_mut().unwrap();
                    for (i, p) in pix.pixels().iter().enumerate() {
                        buffer[i] =
                            (p.red() as u32) << 16 | (p.green() as u32) << 8 | (p.blue() as u32);
                    }
                    buffer.present().unwrap();
                }
            }

            _ => {}
        }
    }
}

/// Run a GUI application with a window.
///
/// `scale` is the HiDPI scale factor (e.g. `2.0` for Retina).
pub fn run_app<A: Application + 'static>(title: &str, width: u32, height: u32, scale: f32, app: A) {
    let el = EventLoop::new().expect("create event loop");
    el.set_control_flow(ControlFlow::Wait);
    let mut inner = AppWindowInner {
        app,
        window: None,
        context: None,
        surface: None,
        pixmap: None,
        scale,
        title: title.to_string(),
        init_width: width,
        init_height: height,
        modifiers: winit::event::Modifiers::default(),
        mouse_pos: (0.0, 0.0),
    };
    el.run_app(&mut inner).expect("run app");
}
