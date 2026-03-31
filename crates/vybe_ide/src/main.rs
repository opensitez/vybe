//! Vybe Skia IDE — pure tiny-skia + winit IDE for VB.NET/JS/Python/C#/Dart projects.

mod app;
mod layout;
mod panels;
mod text;

use std::rc::Rc;

use app::SkiaIde;
use tiny_skia::Pixmap;
use winit::application::ApplicationHandler;
use winit::dpi::{LogicalSize, PhysicalPosition};
use winit::event::{ElementState, MouseButton, MouseScrollDelta, WindowEvent};
use winit::event_loop::{ActiveEventLoop, EventLoop};
use winit::keyboard::{Key, NamedKey};
use winit::window::{Window, WindowAttributes, WindowId};

struct WinitApp {
    window: Option<Rc<Window>>,
    context: Option<softbuffer::Context<Rc<Window>>>,
    surface: Option<softbuffer::Surface<Rc<Window>, Rc<Window>>>,
    ide: SkiaIde,
    needs_redraw: bool,
    cursor_logical: (f32, f32),
    modifiers: winit::event::Modifiers,
}

impl WinitApp {
    fn new() -> Self {
        Self {
            window: None,
            context: None,
            surface: None,
            ide: SkiaIde::new(1.0),
            needs_redraw: true,
            cursor_logical: (0.0, 0.0),
            modifiers: winit::event::Modifiers::default(),
        }
    }

    fn redraw(&mut self) {
        let Some(window) = &self.window else { return };
        let Some(surface) = &mut self.surface else { return };

        let size = window.inner_size();
        let pw = size.width.max(1);
        let ph = size.height.max(1);

        let scale = window.scale_factor() as f32;
        self.ide.scale = scale;
        self.ide.win_w = pw as f32 / scale;
        self.ide.win_h = ph as f32 / scale;

        surface.resize(
            std::num::NonZeroU32::new(pw).unwrap(),
            std::num::NonZeroU32::new(ph).unwrap(),
        ).ok();

        let mut pixmap = Pixmap::new(pw, ph).unwrap();
        self.ide.render(&mut pixmap);

        // Copy pixmap → softbuffer
        if let Ok(mut buf) = surface.buffer_mut() {
            let data = pixmap.data();
            for (i, pixel) in buf.iter_mut().enumerate() {
                let off = i * 4;
                if off + 3 < data.len() {
                    // tiny-skia is premultiplied RGBA, softbuffer wants 0xAARRGGBB or 0x00RRGGBB
                    let r = data[off] as u32;
                    let g = data[off + 1] as u32;
                    let b = data[off + 2] as u32;
                    *pixel = (r << 16) | (g << 8) | b;
                }
            }
            buf.present().ok();
        }

        self.needs_redraw = false;
    }

    fn logical_pos(&self, pos: PhysicalPosition<f64>) -> (f32, f32) {
        let scale = self.window.as_ref().map(|w| w.scale_factor()).unwrap_or(1.0) as f32;
        (pos.x as f32 / scale, pos.y as f32 / scale)
    }

    fn request_redraw(&mut self) {
        self.needs_redraw = true;
        if let Some(w) = &self.window {
            w.request_redraw();
        }
    }
}

impl ApplicationHandler for WinitApp {
    fn resumed(&mut self, event_loop: &ActiveEventLoop) {
        if self.window.is_some() { return; }

        let attrs = WindowAttributes::default()
            .with_title("Vybe IDE")
            .with_inner_size(LogicalSize::new(1200.0, 800.0));

        let window = Rc::new(event_loop.create_window(attrs).unwrap());
        let context = softbuffer::Context::new(window.clone()).unwrap();
        let surface = softbuffer::Surface::new(&context, window.clone()).unwrap();

        self.window = Some(window);
        self.context = Some(context);
        self.surface = Some(surface);
        self.request_redraw();
    }

    fn window_event(&mut self, event_loop: &ActiveEventLoop, _id: WindowId, event: WindowEvent) {
        match event {
            WindowEvent::CloseRequested => {
                event_loop.exit();
            }

            WindowEvent::Resized(_) => {
                self.request_redraw();
            }

            WindowEvent::RedrawRequested => {
                self.redraw();
            }

            WindowEvent::ModifiersChanged(mods) => {
                self.modifiers = mods;
            }

            WindowEvent::MouseInput { state, button: MouseButton::Left, .. } => {
                let (lx, ly) = self.cursor_logical;
                let ctrl = self.modifiers.state().super_key() || self.modifiers.state().control_key();
                match state {
                    ElementState::Pressed => {
                        self.ide.handle_mouse_down(lx, ly, ctrl);
                    }
                    ElementState::Released => {
                        self.ide.handle_mouse_up(lx, ly);
                    }
                }
                self.request_redraw();
            }

            WindowEvent::CursorMoved { position, .. } => {
                let (lx, ly) = self.logical_pos(position);
                self.cursor_logical = (lx, ly);

                // Always update hover state for menus
                self.ide.handle_mouse_hover(lx, ly);

                if self.ide.mouse_down {
                    self.ide.handle_mouse_move(lx, ly);
                }
                self.request_redraw();
            }

            WindowEvent::MouseWheel { delta, .. } => {
                let dy = match delta {
                    MouseScrollDelta::LineDelta(_, y) => y,
                    MouseScrollDelta::PixelDelta(pos) => pos.y as f32 / 20.0,
                };
                let (lx, ly) = self.cursor_logical;
                self.ide.handle_scroll(dy, lx, ly);
                self.request_redraw();
            }

            WindowEvent::KeyboardInput { event, .. } => {
                if event.state != ElementState::Pressed { return; }

                let ctrl = self.modifiers.state().super_key() || self.modifiers.state().control_key();
                let shift = self.modifiers.state().shift_key();

                match &event.logical_key {
                    Key::Named(named) => {
                        let key_str = match named {
                            NamedKey::ArrowLeft => "Left",
                            NamedKey::ArrowRight => "Right",
                            NamedKey::ArrowUp => "Up",
                            NamedKey::ArrowDown => "Down",
                            NamedKey::Home => "Home",
                            NamedKey::End => "End",
                            NamedKey::Enter => "Enter",
                            NamedKey::Backspace => "Backspace",
                            NamedKey::Delete => "Delete",
                            NamedKey::Tab => "Tab",
                            NamedKey::Escape => "Escape",
                            NamedKey::Space => {
                                self.ide.handle_char(' ');
                                self.request_redraw();
                                return;
                            }
                            _ => return,
                        };
                        self.ide.handle_key(key_str);
                        self.request_redraw();
                    }
                    Key::Character(ch) => {
                        if ctrl {
                            // Ctrl+key shortcuts
                            match ch.as_str() {
                                "c" => self.ide.handle_shortcut("copy"),
                                "x" => self.ide.handle_shortcut("cut"),
                                "v" => self.ide.handle_shortcut("paste"),
                                "z" => {
                                    if shift {
                                        self.ide.handle_shortcut("redo");
                                    } else {
                                        self.ide.handle_shortcut("undo");
                                    }
                                }
                                "y" => self.ide.handle_shortcut("redo"),
                                "a" => self.ide.handle_shortcut("select_all"),
                                "s" => self.ide.handle_shortcut("save"),
                                "n" => self.ide.handle_shortcut("new"),
                                "o" => self.ide.handle_shortcut("open"),
                                _ => {}
                            }
                        } else {
                            for c in ch.chars() {
                                self.ide.handle_char(c);
                            }
                        }
                        self.request_redraw();
                    }
                    _ => {}
                }
            }

            _ => {}
        }
    }
}

fn main() {
    let event_loop = EventLoop::new().unwrap();
    let mut app = WinitApp::new();

    // We need to handle mouse clicks properly by tracking cursor position
    // The winit API sends CursorMoved and MouseInput separately
    event_loop.run_app(&mut app).unwrap();
}
