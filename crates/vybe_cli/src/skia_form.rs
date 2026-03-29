//! Tiny-Skia based form renderer — replaces Dioxus webview.
//! Renders controls as pixels using winit + tiny-skia + softbuffer.

use std::num::NonZeroU32;
use std::rc::Rc;
use std::cell::RefCell;
use winit::application::ApplicationHandler;
use winit::event::{WindowEvent, MouseButton, ElementState};
use winit::event_loop::{EventLoop, ActiveEventLoop, ControlFlow};
use winit::window::{Window, WindowId, WindowAttributes};

use tiny_skia::*;

/// A rendered control — position, size, text, type.
struct RenderedControl {
    name: String,
    type_name: String,
    x: i32,
    y: i32,
    w: i32,
    h: i32,
    text: String,
}

/// Colors
fn bg_color() -> Color { Color::from_rgba8(240, 240, 240, 255) }
fn btn_color() -> Color { Color::from_rgba8(225, 225, 225, 255) }
fn btn_border() -> Color { Color::from_rgba8(173, 173, 173, 255) }
fn text_color() -> Color { Color::from_rgba8(0, 0, 0, 255) }
fn input_bg() -> Color { Color::from_rgba8(255, 255, 255, 255) }
fn input_border() -> Color { Color::from_rgba8(122, 122, 122, 255) }
fn panel_border() -> Color { Color::from_rgba8(200, 200, 200, 255) }
fn progress_bg() -> Color { Color::from_rgba8(230, 230, 230, 255) }
fn progress_fg() -> Color { Color::from_rgba8(6, 176, 37, 255) }
fn grid_header() -> Color { Color::from_rgba8(240, 240, 240, 255) }
fn grid_line() -> Color { Color::from_rgba8(200, 200, 200, 255) }

/// Simple font rendering using fontdue
struct FontRenderer {
    font: fontdue::Font,
}

impl FontRenderer {
    fn new() -> Self {
        // Try system fonts at runtime
        let font_paths = [
            "/System/Library/Fonts/Helvetica.ttc",
            "/System/Library/Fonts/SFNSText.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/TTF/DejaVuSans.ttf",
            "C:\\Windows\\Fonts\\segoeui.ttf",
            "C:\\Windows\\Fonts\\arial.ttf",
        ];
        for path in &font_paths {
            if let Ok(data) = std::fs::read(path) {
                let settings = fontdue::FontSettings {
                    collection_index: 0,
                    scale: 40.0,
                    load_substitutions: true,
                };
                if let Ok(font) = fontdue::Font::from_bytes(data, settings) {
                    return FontRenderer { font };
                }
            }
        }
        panic!("No system font found");
    }

    fn draw_text(&self, pixmap: &mut Pixmap, text: &str, x: f32, y: f32, size: f32, color: Color) {
        let mut cx = x;
        for ch in text.chars() {
            let (metrics, bitmap) = self.font.rasterize(ch, size);
            if bitmap.is_empty() {
                cx += metrics.advance_width;
                continue;
            }
            let bx = cx as i32 + metrics.xmin;
            let by = y as i32 - metrics.ymin - metrics.height as i32 + (size * 0.8) as i32;
            let pw = pixmap.width() as i32;
            let ph = pixmap.height() as i32;
            let pixels = pixmap.pixels_mut();
            for py in 0..metrics.height {
                for px in 0..metrics.width {
                    let alpha = bitmap[py * metrics.width + px];
                    if alpha == 0 { continue; }
                    let dx = bx + px as i32;
                    let dy = by + py as i32;
                    if dx >= 0 && dx < pw && dy >= 0 && dy < ph {
                        let idx = (dy * pw + dx) as usize;
                        let existing = pixels[idx];
                        let a = alpha as f32 / 255.0;
                        let r = (color.red() * a + existing.red() as f32 / 255.0 * (1.0 - a)).min(1.0);
                        let g = (color.green() * a + existing.green() as f32 / 255.0 * (1.0 - a)).min(1.0);
                        let b = (color.blue() * a + existing.blue() as f32 / 255.0 * (1.0 - a)).min(1.0);
                        pixels[idx] = ColorU8::from_rgba(
                            (r * 255.0) as u8, (g * 255.0) as u8, (b * 255.0) as u8, 255
                        ).premultiply();
                    }
                }
            }
            cx += metrics.advance_width;
        }
    }

    fn text_width(&self, text: &str, size: f32) -> f32 {
        text.chars().map(|ch| {
            let (metrics, _) = self.font.rasterize(ch, size);
            metrics.advance_width
        }).sum()
    }
}

fn fill_rect(pixmap: &mut Pixmap, x: i32, y: i32, w: i32, h: i32, color: Color) {
    let rect = Rect::from_xywh(x as f32, y as f32, w as f32, h as f32);
    if let Some(rect) = rect {
        let mut paint = Paint::default();
        paint.set_color(color);
        pixmap.fill_rect(rect, &paint, Transform::identity(), None);
    }
}

fn stroke_rect(pixmap: &mut Pixmap, x: i32, y: i32, w: i32, h: i32, color: Color) {
    let path = {
        let mut pb = PathBuilder::new();
        pb.move_to(x as f32, y as f32);
        pb.line_to((x + w) as f32, y as f32);
        pb.line_to((x + w) as f32, (y + h) as f32);
        pb.line_to(x as f32, (y + h) as f32);
        pb.close();
        pb.finish().unwrap()
    };
    let mut paint = Paint::default();
    paint.set_color(color);
    let stroke = Stroke { width: 1.0, ..Stroke::default() };
    pixmap.stroke_path(&path, &paint, &stroke, Transform::identity(), None);
}

fn render_control_scaled(pixmap: &mut Pixmap, font: &FontRenderer, ctrl: &RenderedControl, scale: f32) {
    render_control_impl(pixmap, font, ctrl, 13.0 * scale, 11.0 * scale);
}

fn render_control(pixmap: &mut Pixmap, font: &FontRenderer, ctrl: &RenderedControl) {
    render_control_impl(pixmap, font, ctrl, 13.0, 11.0);
}

fn render_control_impl(pixmap: &mut Pixmap, font: &FontRenderer, ctrl: &RenderedControl, font_size: f32, small_size: f32) {
    let (x, y, w, h) = (ctrl.x, ctrl.y, ctrl.w, ctrl.h);
    let type_lower = ctrl.type_name.to_lowercase();

    match type_lower.as_str() {
        "button" => {
            fill_rect(pixmap, x, y, w, h, btn_color());
            stroke_rect(pixmap, x, y, w, h, btn_border());
            // Center text
            let tw = font.text_width(&ctrl.text, font_size);
            let tx = x as f32 + (w as f32 - tw) / 2.0;
            let ty = y as f32 + (h as f32) / 2.0 - 2.0;
            font.draw_text(pixmap, &ctrl.text, tx, ty, font_size, text_color());
        }
        "label" | "linklabel" => {
            font.draw_text(pixmap, &ctrl.text, x as f32 + 2.0, y as f32 + 4.0, font_size, text_color());
        }
        "textbox" | "maskedtextbox" | "richtextbox" => {
            fill_rect(pixmap, x, y, w, h, input_bg());
            stroke_rect(pixmap, x, y, w, h, input_border());
            font.draw_text(pixmap, &ctrl.text, x as f32 + 4.0, y as f32 + 4.0, font_size, text_color());
        }
        "checkbox" => {
            // Box + text
            fill_rect(pixmap, x, y, 16, 16, input_bg());
            stroke_rect(pixmap, x, y, 16, 16, input_border());
            font.draw_text(pixmap, &ctrl.text, x as f32 + 20.0, y as f32 + 2.0, font_size, text_color());
        }
        "radiobutton" => {
            // Circle placeholder + text
            stroke_rect(pixmap, x, y, 16, 16, input_border());
            font.draw_text(pixmap, &ctrl.text, x as f32 + 20.0, y as f32 + 2.0, font_size, text_color());
        }
        "combobox" => {
            fill_rect(pixmap, x, y, w, h, input_bg());
            stroke_rect(pixmap, x, y, w, h, input_border());
            font.draw_text(pixmap, &ctrl.text, x as f32 + 4.0, y as f32 + 4.0, font_size, text_color());
            // Dropdown arrow
            font.draw_text(pixmap, "v", (x + w - 16) as f32, y as f32 + 4.0, font_size, text_color());
        }
        "listbox" => {
            fill_rect(pixmap, x, y, w, h, input_bg());
            stroke_rect(pixmap, x, y, w, h, input_border());
        }
        "panel" | "groupbox" | "frame" | "splitcontainer" | "flowlayoutpanel" | "tablelayoutpanel" => {
            stroke_rect(pixmap, x, y, w, h, panel_border());
            if !ctrl.text.is_empty() {
                font.draw_text(pixmap, &ctrl.text, x as f32 + 8.0, y as f32 - 2.0, small_size, text_color());
            }
        }
        "tabcontrol" => {
            stroke_rect(pixmap, x, y, w, h, panel_border());
            fill_rect(pixmap, x, y, w, 24, btn_color());
            font.draw_text(pixmap, "Tab1", x as f32 + 8.0, y as f32 + 5.0, small_size, text_color());
        }
        "progressbar" => {
            fill_rect(pixmap, x, y, w, h, progress_bg());
            stroke_rect(pixmap, x, y, w, h, panel_border());
            // 50% fill
            fill_rect(pixmap, x + 1, y + 1, w / 2, h - 2, progress_fg());
        }
        "trackbar" => {
            let mid_y = y + h / 2;
            fill_rect(pixmap, x, mid_y - 2, w, 4, progress_bg());
            stroke_rect(pixmap, x, mid_y - 2, w, 4, panel_border());
            // Thumb
            fill_rect(pixmap, x + w / 2 - 4, mid_y - 8, 8, 16, btn_color());
            stroke_rect(pixmap, x + w / 2 - 4, mid_y - 8, 8, 16, btn_border());
        }
        "datagridview" | "listview" => {
            let header_h = (font_size * 1.8) as i32;
            let row_h = (font_size * 1.6) as i32;
            fill_rect(pixmap, x, y, w, h, input_bg());
            stroke_rect(pixmap, x, y, w, h, input_border());
            // Header bar
            fill_rect(pixmap, x + 1, y + 1, w - 2, header_h, grid_header());
            stroke_rect(pixmap, x + 1, y + 1, w - 2, header_h, grid_line());
            // Column headers placeholder
            let col_w = w / 3;
            for col in 0..3 {
                let cx = x + 1 + col * col_w;
                font.draw_text(pixmap, &format!("Column {}", col + 1), cx as f32 + 4.0, y as f32 + 4.0, small_size, text_color());
                // Vertical separator
                if col > 0 {
                    fill_rect(pixmap, cx, y + 1, 1, h - 2, grid_line());
                }
            }
            // Row lines
            let mut row_y = y + 1 + header_h;
            let mut row_num = 0;
            while row_y < y + h - 1 {
                fill_rect(pixmap, x + 1, row_y, w - 2, 1, grid_line());
                // Alternating row background
                if row_num % 2 == 1 && row_y + row_h < y + h {
                    fill_rect(pixmap, x + 1, row_y + 1, w - 2, row_h - 1,
                        Color::from_rgba8(245, 245, 250, 255));
                }
                row_y += row_h;
                row_num += 1;
            }
        }
        "numericupdown" | "datetimepicker" => {
            fill_rect(pixmap, x, y, w, h, input_bg());
            stroke_rect(pixmap, x, y, w, h, input_border());
            font.draw_text(pixmap, &ctrl.text, x as f32 + 4.0, y as f32 + 4.0, font_size, text_color());
            // Up/down buttons
            fill_rect(pixmap, x + w - 18, y, 18, h, btn_color());
            stroke_rect(pixmap, x + w - 18, y, 18, h, btn_border());
        }
        "bindingnavigator" => {
            let btn_w = (font_size * 2.0) as i32;
            let gap = 2;
            fill_rect(pixmap, x, y, w, h, btn_color());
            stroke_rect(pixmap, x, y, w, h, panel_border());
            // Navigation buttons: |< < record counter > >| + -
            let buttons = ["|<", "<", "", ">", ">|", "+", "-"];
            let mut bx = x + 2;
            for (i, label) in buttons.iter().enumerate() {
                if i == 2 {
                    // Record counter textbox
                    let counter_w = btn_w * 2;
                    fill_rect(pixmap, bx, y + 2, counter_w, h - 4, input_bg());
                    stroke_rect(pixmap, bx, y + 2, counter_w, h - 4, input_border());
                    font.draw_text(pixmap, "0 of 0", bx as f32 + 4.0, y as f32 + 3.0, small_size, text_color());
                    bx += counter_w + gap;
                } else {
                    fill_rect(pixmap, bx, y + 2, btn_w, h - 4, Color::from_rgba8(235, 235, 235, 255));
                    stroke_rect(pixmap, bx, y + 2, btn_w, h - 4, panel_border());
                    let tw = font.text_width(label, small_size);
                    font.draw_text(pixmap, label, bx as f32 + (btn_w as f32 - tw) / 2.0, y as f32 + 3.0, small_size, text_color());
                    bx += btn_w + gap;
                }
            }
        }
        "menustrip" | "toolstrip" | "statusstrip" | "contextmenustrip" => {
            fill_rect(pixmap, x, y, w, h, btn_color());
            stroke_rect(pixmap, x, y, w, h, panel_border());
            font.draw_text(pixmap, &ctrl.text, x as f32 + 4.0, y as f32 + 2.0, small_size, text_color());
        }
        "picturebox" => {
            fill_rect(pixmap, x, y, w, h, input_bg());
            stroke_rect(pixmap, x, y, w, h, panel_border());
            font.draw_text(pixmap, "[Image]", x as f32 + 4.0, (y + h / 2 - 6) as f32, small_size, panel_border());
        }
        "webbrowser" => {
            fill_rect(pixmap, x, y, w, h, input_bg());
            stroke_rect(pixmap, x, y, w, h, panel_border());
            font.draw_text(pixmap, "[WebBrowser]", x as f32 + 4.0, (y + h / 2 - 6) as f32, small_size, panel_border());
        }
        "treeview" => {
            fill_rect(pixmap, x, y, w, h, input_bg());
            stroke_rect(pixmap, x, y, w, h, input_border());
            font.draw_text(pixmap, "- Node 1", x as f32 + 4.0, y as f32 + 4.0, small_size, text_color());
        }
        "monthcalendar" => {
            fill_rect(pixmap, x, y, w, h, input_bg());
            stroke_rect(pixmap, x, y, w, h, input_border());
            font.draw_text(pixmap, "[Calendar]", x as f32 + 4.0, y as f32 + 4.0, small_size, text_color());
        }
        "hscrollbar" | "vscrollbar" => {
            fill_rect(pixmap, x, y, w, h, progress_bg());
            stroke_rect(pixmap, x, y, w, h, panel_border());
        }
        _ => {
            // Non-visual or unknown — skip rendering (BindingSource, DataSet, etc.)
            if w > 32 || h > 32 {
                stroke_rect(pixmap, x, y, w, h, panel_border());
                font.draw_text(pixmap, &ctrl.text, x as f32 + 4.0, y as f32 + 4.0, small_size, text_color());
            }
        }
    }
}

struct FormApp {
    window: Option<Rc<Window>>,
    surface: Option<softbuffer::Surface<Rc<Window>, Rc<Window>>>,
    controls: Vec<RenderedControl>,
    form_width: u32,
    form_height: u32,
    font: FontRenderer,
    vm: Rc<RefCell<vybe_bytecode::VM>>,
    queue: Rc<RefCell<vybe_host::SideEffectQueue>>,
    form_obj_key: String,
    needs_redraw: bool,
    last_cursor: (f64, f64),
}

impl FormApp {
    fn hit_test(&self, mx: f32, my: f32) -> Option<&RenderedControl> {
        self.controls.iter().find(|c| {
            mx >= c.x as f32 && mx <= (c.x + c.w) as f32 &&
            my >= c.y as f32 && my <= (c.y + c.h) as f32
        })
    }

    /// Fire the Form_Load event — called when the form is first shown.
    fn fire_load_event(&mut self) {
        // Look for a Load handler on the form name
        let form_name = self.form_obj_key.replace("__", "");
        let callback = {
            let q = self.queue.borrow();
            // Try common Load event registrations
            q.get_event_handler(&"form1", "Load").cloned()
                .or_else(|| q.get_event_handler(&"me", "Load").cloned())
        };
        if let Some(cb) = callback {
            eprintln!("[LOAD] Firing Form_Load event");
            let mut vm = self.vm.borrow_mut();
            let me = vm.globals.get("__f").cloned()
                .unwrap_or(vybe_bytecode::Value::Null);
            let arity = match &cb {
                vybe_bytecode::Value::Object(obj) => {
                    match &obj.borrow().kind {
                        vybe_bytecode::value::ObjectKind::Function(f) => f.arity as usize,
                        _ => 0,
                    }
                }
                _ => 0,
            };
            let result = match arity {
                0 => vm.invoke(&cb, &[]),
                1 => vm.invoke(&cb, &[me]),
                _ => vm.invoke(&cb, &[me, vybe_bytecode::Value::Null, vybe_bytecode::Value::Null]),
            };
            if let Err(e) = result {
                eprintln!("[LOAD] Error: {e}");
            }
            // Drain console output
            let effects = self.queue.borrow_mut().drain();
            for effect in effects {
                if let vybe_host::SideEffect::ConsoleOutput(msg) = effect {
                    print!("{msg}");
                }
            }
            drop(vm);
            self.read_controls_from_vm();
            self.needs_redraw = true;
        }
    }

    fn handle_click(&mut self, control_name: &str) {
        let callback = {
            let q = self.queue.borrow();
            q.get_event_handler(&control_name.to_lowercase(), "Click").cloned()
        };
        if let Some(cb) = callback {
            let mut vm = self.vm.borrow_mut();
            let me = vm.globals.get("__f").cloned()
                .unwrap_or(vybe_bytecode::Value::Null);
            let arity = match &cb {
                vybe_bytecode::Value::Object(obj) => {
                    match &obj.borrow().kind {
                        vybe_bytecode::value::ObjectKind::Function(f) => f.arity as usize,
                        _ => 0,
                    }
                }
                _ => 0,
            };
            let sender = vybe_bytecode::Value::String(Rc::from(control_name));
            let result = match arity {
                0 => vm.invoke(&cb, &[]),
                1 => vm.invoke(&cb, &[me]),
                2 => vm.invoke(&cb, &[me, sender]),
                _ => vm.invoke(&cb, &[me, sender, vybe_bytecode::Value::Null]),
            };
            if let Err(e) = result {
                eprintln!("Event handler error: {e}");
            }

            // Drain side effects (console output etc.) but don't use for rendering
            let effects = self.queue.borrow_mut().drain();
            for effect in effects {
                if let vybe_host::SideEffect::ConsoleOutput(msg) = effect {
                    print!("{msg}");
                }
            }

            drop(vm);

            // Read control state directly from VM objects — single source of truth
            self.read_controls_from_vm();
            self.needs_redraw = true;
        }
    }

    /// Read all control text/properties directly from VM objects.
    /// No side effects, no form model copy — the VM IS the model.
    fn read_controls_from_vm(&mut self) {
        let vm = self.vm.borrow();
        let has_f = vm.globals.contains_key("__f");
        if !has_f {
            eprintln!("[READ] __f not found!");
            return;
        }
        if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
            let fo = form_obj.borrow();
            let prop_keys: Vec<_> = fo.properties.keys().take(15).cloned().collect();
            eprintln!("[READ] __f has {} props: {:?}", fo.properties.len(), prop_keys);
            for ctrl in &mut self.controls {
                let ctrl_lower = ctrl.name.to_lowercase();
                if let Some(vybe_bytecode::Value::Object(co)) = fo.properties.get(&ctrl_lower) {
                    let c = co.borrow();
                    let text_val = c.properties.get("text").cloned();
                    eprintln!("[READ] {}: text={:?}", ctrl.name, text_val);
                    if let Some(vybe_bytecode::Value::String(s)) = text_val {
                        ctrl.text = s.to_string();
                    } else if let Some(v) = c.properties.get("text") {
                        ctrl.text = format!("{}", v);
                    }
                } else {
                    eprintln!("[READ] {} not found on __f", ctrl.name);
                }
            }
        }
    }

    fn render(&self, pixmap: &mut Pixmap) {
        // Clear background
        pixmap.fill(bg_color());

        // Render all controls
        for ctrl in &self.controls {
            render_control(pixmap, &self.font, ctrl);
        }
    }
}

impl ApplicationHandler for FormApp {
    fn resumed(&mut self, event_loop: &ActiveEventLoop) {
        if self.window.is_none() {
            let attrs = WindowAttributes::default()
                .with_title("Form")
                .with_inner_size(winit::dpi::LogicalSize::new(self.form_width, self.form_height));
            let window = Rc::new(event_loop.create_window(attrs).unwrap());
            let context = softbuffer::Context::new(window.clone()).unwrap();
            let surface = softbuffer::Surface::new(&context, window.clone()).unwrap();
            self.window = Some(window);
            self.surface = Some(surface);
            self.needs_redraw = true;

            // Fire Form_Load event — .NET fires Load when the form is first shown
            self.fire_load_event();
        }
    }

    fn window_event(&mut self, event_loop: &ActiveEventLoop, _id: WindowId, event: WindowEvent) {
        match event {
            WindowEvent::CloseRequested => {
                event_loop.exit();
            }
            WindowEvent::RedrawRequested => {
                let (scale_factor, size) = if let Some(w) = &self.window {
                    (w.scale_factor(), w.inner_size())
                } else { return };
                let w = size.width.max(1);
                let h = size.height.max(1);

                if let Some(surface) = &mut self.surface {
                    surface.resize(NonZeroU32::new(w).unwrap(), NonZeroU32::new(h).unwrap()).ok();

                    let mut pixmap = Pixmap::new(w, h).unwrap();
                    let s = scale_factor as f32;
                    // Render inline — scale all controls by HiDPI factor
                    pixmap.fill(bg_color());
                    for ctrl in &self.controls {
                        let scaled = RenderedControl {
                            name: ctrl.name.clone(),
                            type_name: ctrl.type_name.clone(),
                            x: (ctrl.x as f32 * s) as i32,
                            y: (ctrl.y as f32 * s) as i32,
                            w: (ctrl.w as f32 * s) as i32,
                            h: (ctrl.h as f32 * s) as i32,
                            text: ctrl.text.clone(),
                        };
                        render_control_scaled(&mut pixmap, &self.font, &scaled, s);
                    }

                    // Copy to surface
                    let mut buffer = surface.buffer_mut().unwrap();
                    for (i, pixel) in pixmap.pixels().iter().enumerate() {
                        let r = pixel.red() as u32;
                        let g = pixel.green() as u32;
                        let b = pixel.blue() as u32;
                        buffer[i] = (r << 16) | (g << 8) | b;
                    }
                    buffer.present().ok();
                }
                self.needs_redraw = false;
            }
            WindowEvent::MouseInput { state, button, .. } => {
                let scale = self.window.as_ref().map(|w| w.scale_factor()).unwrap_or(1.0);
                let (lx, ly) = (self.last_cursor.0 / scale, self.last_cursor.1 / scale);
                eprintln!("[MOUSE] {:?} {:?} at logical ({:.0}, {:.0}) physical ({:.0}, {:.0}) scale={}", state, button, lx, ly, self.last_cursor.0, self.last_cursor.1, scale);
                if state != ElementState::Pressed || button != MouseButton::Left {
                    // Only handle left press
                } else {
                let (mx, my) = (lx, ly);
                // Hit test against controls (using logical coords)
                let clicked = self.controls.iter().find(|c| {
                    mx >= c.x as f64 && mx <= (c.x + c.w) as f64 &&
                    my >= c.y as f64 && my <= (c.y + c.h) as f64
                }).map(|c| c.name.clone());
                if let Some(name) = clicked {
                    eprintln!("[SKIA-CLICK] {}", name);
                    self.handle_click(&name);
                    // Request redraw to show updated state
                    if let Some(window) = &self.window {
                        window.request_redraw();
                    }
                }
                } // end left press
            }
            WindowEvent::CursorMoved { position, .. } => {
                self.last_cursor = (position.x, position.y);
            }
            _ => {}
        }

        if self.needs_redraw {
            if let Some(window) = &self.window {
                window.request_redraw();
            }
        }
    }
}

/// Launch a form using tiny-skia rendering.
/// Register native dialog host functions using rfd.
fn register_dialog_fns(vm: &mut vybe_bytecode::VM) {
    use vybe_bytecode::Value;
    use vybe_bytecode::value::{Object, ObjectKind};

    // ShowDialog on dialog objects — returns DialogResult.OK (1) or Cancel (0)
    vm.register_host_fn("vybe:gui", "__dlg_show", Box::new(|args: &[Value]| {
        let dialog_type = if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            o.properties.get("__control_type").map(|v| format!("{}", v)).unwrap_or_default()
        } else { String::new() };

        match dialog_type.as_str() {
            "OpenFileDialog" => {
                let result = rfd::FileDialog::new()
                    .set_title("Open File")
                    .pick_file();
                if let Some(path) = result {
                    // Store filename on the dialog object
                    if let Some(Value::Object(obj)) = args.first() {
                        obj.borrow_mut().properties.insert("filename".into(),
                            Value::String(Rc::from(path.to_string_lossy().as_ref())));
                    }
                    Value::I32(1) // DialogResult.OK
                } else {
                    Value::I32(0) // DialogResult.Cancel
                }
            }
            "SaveFileDialog" => {
                let result = rfd::FileDialog::new()
                    .set_title("Save File")
                    .save_file();
                if let Some(path) = result {
                    if let Some(Value::Object(obj)) = args.first() {
                        obj.borrow_mut().properties.insert("filename".into(),
                            Value::String(Rc::from(path.to_string_lossy().as_ref())));
                    }
                    Value::I32(1)
                } else {
                    Value::I32(0)
                }
            }
            "FolderBrowserDialog" => {
                let result = rfd::FileDialog::new()
                    .set_title("Select Folder")
                    .pick_folder();
                if let Some(path) = result {
                    if let Some(Value::Object(obj)) = args.first() {
                        obj.borrow_mut().properties.insert("selectedpath".into(),
                            Value::String(Rc::from(path.to_string_lossy().as_ref())));
                    }
                    Value::I32(1)
                } else {
                    Value::I32(0)
                }
            }
            "ColorDialog" | "FontDialog" => {
                // Stub — return OK for now
                Value::I32(1)
            }
            _ => Value::I32(0),
        }
    }));

    // MessageBox.Show — native dialog
    vm.register_host_fn("vybe:gui", "msgBox", Box::new(|args: &[Value]| {
        let text = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let title = args.get(1).map(|v| format!("{}", v)).unwrap_or_else(|| "Message".into());
        rfd::MessageDialog::new()
            .set_title(&title)
            .set_description(&text)
            .set_level(rfd::MessageLevel::Info)
            .show();
        Value::Null
    }));

    // InputBox — native text input (rfd doesn't have this, use a simple stub)
    vm.register_host_fn("vybe:gui", "inputBox", Box::new(|args: &[Value]| {
        let _prompt = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let _title = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let default = args.get(2).map(|v| format!("{}", v)).unwrap_or_default();
        // rfd doesn't support text input — return default for now
        Value::String(Rc::from(default.as_str()))
    }));

    // Make ShowDialog available on dialog objects
    let dlg_show_idx = *vm.host_registry.get(&("vybe:gui".into(), "__dlg_show".into())).unwrap();
    let dlg_show_ref = {
        let mut o = Object::new();
        o.kind = ObjectKind::HostFunction(dlg_show_idx);
        Value::Object(Rc::new(RefCell::new(o)))
    };
    // Store for later attachment to dialog objects
    vm.globals.insert("__dlg_show_ref".into(), dlg_show_ref);
}

pub fn launch_skia_form(
    mut vm: vybe_bytecode::VM,
    queue: Rc<RefCell<vybe_host::SideEffectQueue>>,
    form: &vybe_forms::Form,
    title: &str,
) {
    // Register native dialog functions (rfd)
    register_dialog_fns(&mut vm);
    let controls: Vec<RenderedControl> = form.controls.iter()
        .filter(|c| !c.control_type.is_non_visual())
        .map(|ctrl| {
            RenderedControl {
                name: ctrl.name.clone(),
                type_name: format!("{:?}", ctrl.control_type),
                x: ctrl.bounds.x,
                y: ctrl.bounds.y,
                w: ctrl.bounds.width,
                h: ctrl.bounds.height,
                text: ctrl.properties.get_string("Text").unwrap_or_default().to_string(),
            }
        })
        .collect();

    let event_loop = EventLoop::new().unwrap();
    event_loop.set_control_flow(ControlFlow::Wait);

    let mut app = FormApp {
        window: None,
        surface: None,
        controls,
        form_width: form.width as u32,
        form_height: form.height as u32,
        font: FontRenderer::new(),
        vm: Rc::new(RefCell::new(vm)),
        queue,
        form_obj_key: "__f".into(),
        needs_redraw: true,
        last_cursor: (0.0, 0.0),
    };

    // Store title for window creation
    // (winit requires title at creation, handled in resumed)

    event_loop.run_app(&mut app).ok();
}
