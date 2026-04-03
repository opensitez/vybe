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

/// A rendered control — position, size, text, type + optional data binding state.
struct RenderedControl {
    name: String,
    type_name: String,
    x: i32,
    y: i32,
    w: i32,
    h: i32,
    text: String,
    // Data binding state (populated for data-bound controls)
    grid_columns: Vec<String>,
    grid_rows: Vec<Vec<String>>,
    nav_position: i32,
    nav_count: i32,
}

/// A single data binding: control.property ← bindingSource[column]
#[derive(Clone, Debug)]
struct DataBindingEntry {
    control_name: String,
    property: String,
    source_name: String,
    column: String,
}

/// Info about a BindingSource extracted from the form model.
#[derive(Clone, Debug)]
struct BindingSourceInfo {
    name: String,
    data_adapter_name: String,
    data_member: String,
}

/// Info about a BindingNavigator's linked BindingSource.
#[derive(Clone, Debug)]
struct NavigatorInfo {
    navigator_name: String,
    binding_source_name: String,
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

            let num_cols = if ctrl.grid_columns.is_empty() { 3 } else { ctrl.grid_columns.len() };
            let col_w = if num_cols > 0 { (w - 2) / num_cols as i32 } else { w / 3 };

            // Column headers
            for col in 0..num_cols {
                let cx = x + 1 + col as i32 * col_w;
                let header_text = if col < ctrl.grid_columns.len() {
                    ctrl.grid_columns[col].as_str()
                } else {
                    ""
                };
                font.draw_text(pixmap, header_text, cx as f32 + 4.0, y as f32 + 4.0, small_size, text_color());
                if col > 0 {
                    fill_rect(pixmap, cx, y + 1, 1, h - 2, grid_line());
                }
            }

            // Data rows
            let mut row_y = y + 1 + header_h;
            let mut row_num = 0;
            let max_visible = ((h - header_h - 2) / row_h).max(0) as usize;
            while row_num < max_visible {
                fill_rect(pixmap, x + 1, row_y, w - 2, 1, grid_line());
                if row_num % 2 == 1 && row_y + row_h < y + h {
                    fill_rect(pixmap, x + 1, row_y + 1, w - 2, row_h - 1,
                        Color::from_rgba8(245, 245, 250, 255));
                }
                // Highlight current row
                if row_num < ctrl.grid_rows.len() && row_num == ctrl.nav_position as usize {
                    fill_rect(pixmap, x + 1, row_y + 1, w - 2, row_h - 1,
                        Color::from_rgba8(51, 153, 255, 60));
                }
                // Draw cell values
                if row_num < ctrl.grid_rows.len() {
                    let row_data = &ctrl.grid_rows[row_num];
                    for col in 0..num_cols.min(row_data.len()) {
                        let cx = x + 1 + col as i32 * col_w;
                        let cell_text = &row_data[col];
                        // Clip text to column width
                        let max_chars = (col_w as f32 / (small_size * 0.6)) as usize;
                        let display = if cell_text.len() > max_chars {
                            &cell_text[..max_chars.max(1)]
                        } else {
                            cell_text.as_str()
                        };
                        font.draw_text(pixmap, display, cx as f32 + 4.0, row_y as f32 + 3.0, small_size, text_color());
                    }
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
                    let counter_text = format!("{} of {}", ctrl.nav_position + 1, ctrl.nav_count);
                    font.draw_text(pixmap, &counter_text, bx as f32 + 4.0, y as f32 + 3.0, small_size, text_color());
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
    // Data binding state
    data_bindings: Vec<DataBindingEntry>,
    binding_sources: Vec<BindingSourceInfo>,
    navigators: Vec<NavigatorInfo>,
    /// column_name → Vec<cell_values> per BindingSource, keyed by bs name
    data_store: std::collections::HashMap<String, DataStore>,
}

/// Populated data for a BindingSource.
struct DataStore {
    columns: Vec<String>,
    rows: Vec<std::collections::HashMap<String, String>>,
    position: i32,
}

impl FormApp {
    #[allow(dead_code)]
    fn hit_test(&self, mx: f32, my: f32) -> Option<&RenderedControl> {
        self.controls.iter().find(|c| {
            mx >= c.x as f32 && mx <= (c.x + c.w) as f32 &&
            my >= c.y as f32 && my <= (c.y + c.h) as f32
        })
    }

    /// Fire the Form_Load event — called when the form is first shown.
    fn fire_load_event(&mut self) {
        // Look for a Load handler on the form name
        let _form_name = self.form_obj_key.replace("__", "");
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

    /// Initialize data bindings: connect to DB, fill data, sync controls.
    /// Called once after fire_load_event when the form has BindingSources with DataAdapters.
    fn init_data_bindings(&mut self) {
        if self.binding_sources.is_empty() {
            return;
        }
        eprintln!("[DATA] Initializing {} binding source(s), {} binding(s), {} navigator(s)",
            self.binding_sources.len(), self.data_bindings.len(), self.navigators.len());

        let bs_infos: Vec<_> = self.binding_sources.clone();
        for bs_info in &bs_infos {
            // Get the DataAdapter's ConnectionString from the VM
            let conn_str = {
                let vm = self.vm.borrow();
                if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
                    let fo = form_obj.borrow();
                    // Get the BindingSource object
                    if let Some(vybe_bytecode::Value::Object(bs_obj)) = fo.properties.get(&bs_info.name.to_lowercase()) {
                        let bs = bs_obj.borrow();
                        // DataSource is a reference to the DataAdapter object
                        if let Some(vybe_bytecode::Value::Object(da_obj)) = bs.properties.get("datasource") {
                            let da = da_obj.borrow();
                            da.properties.get("connectionstring")
                                .map(|v| format!("{}", v))
                                .unwrap_or_default()
                        } else {
                            // Try from the adapter name directly on the form
                            if let Some(vybe_bytecode::Value::Object(da_obj)) = fo.properties.get(&bs_info.data_adapter_name.to_lowercase()) {
                                let da = da_obj.borrow();
                                da.properties.get("connectionstring")
                                    .map(|v| format!("{}", v))
                                    .unwrap_or_default()
                            } else {
                                String::new()
                            }
                        }
                    } else {
                        String::new()
                    }
                } else {
                    String::new()
                }
            };

            if conn_str.is_empty() {
                eprintln!("[DATA] No connection string found for BindingSource '{}'", bs_info.name);
                continue;
            }

            let sql = format!("SELECT * FROM {}", bs_info.data_member);
            eprintln!("[DATA] Querying: {} (conn={})", sql, conn_str);

            match vybe_host::modules::database::query_rows(&conn_str, &sql) {
                Ok((columns, rows)) => {
                    eprintln!("[DATA] Got {} rows, {} columns: {:?}", rows.len(), columns.len(), columns);
                    let store = DataStore {
                        columns: columns.clone(),
                        rows: rows.clone(),
                        position: if rows.is_empty() { -1 } else { 0 },
                    };
                    self.data_store.insert(bs_info.name.to_lowercase(), store);

                    // Sync bound controls with position 0
                    self.sync_bound_controls(&bs_info.name);
                }
                Err(e) => {
                    eprintln!("[DATA] Query error for '{}': {}", bs_info.name, e);
                    // Store empty data so rendering still works
                    self.data_store.insert(bs_info.name.to_lowercase(), DataStore {
                        columns: Vec::new(),
                        rows: Vec::new(),
                        position: -1,
                    });
                }
            }
        }

        self.update_data_controls();
        self.needs_redraw = true;
    }

    /// Sync bound TextBox/control properties from the current row in BindingSource.
    fn sync_bound_controls(&mut self, bs_name: &str) {
        let bs_lower = bs_name.to_lowercase();
        let store = match self.data_store.get(&bs_lower) {
            Some(s) => s,
            None => return,
        };
        if store.position < 0 || store.position as usize >= store.rows.len() {
            return;
        }
        let row = &store.rows[store.position as usize];

        // Update VM objects for bound controls
        let vm = self.vm.borrow_mut();
        if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
            let fo = form_obj.borrow();
            for binding in &self.data_bindings {
                if !binding.source_name.eq_ignore_ascii_case(bs_name) {
                    continue;
                }
                let col_key = row.keys()
                    .find(|k| k.eq_ignore_ascii_case(&binding.column))
                    .cloned();
                let value = col_key.and_then(|k| row.get(&k)).cloned().unwrap_or_default();

                // Update the control property on the VM object
                let ctrl_lower = binding.control_name.to_lowercase();
                if let Some(vybe_bytecode::Value::Object(ctrl_obj)) = fo.properties.get(&ctrl_lower) {
                    let prop_lower = binding.property.to_lowercase();
                    ctrl_obj.borrow_mut().properties.insert(
                        prop_lower,
                        vybe_bytecode::Value::String(Rc::from(value.as_str())),
                    );
                }
            }
        }
        drop(vm);

        // Update displayed text from VM
        self.read_controls_from_vm();
    }

    /// Update DataGridView and BindingNavigator controls with current data.
    fn update_data_controls(&mut self) {
        // Collect grid → BindingSource mappings first to avoid borrow conflicts
        let grid_bs_map: Vec<(String, String)> = self.controls.iter()
            .filter(|c| {
                let t = c.type_name.to_lowercase();
                t == "datagridview" || t == "listview"
            })
            .filter_map(|c| {
                self.find_grid_binding_source(&c.name).map(|bs| (c.name.clone(), bs))
            })
            .collect();

        for ctrl in &mut self.controls {
            let type_lower = ctrl.type_name.to_lowercase();
            if type_lower == "datagridview" || type_lower == "listview" {
                if let Some((_, bs_name)) = grid_bs_map.iter().find(|(g, _)| g.eq_ignore_ascii_case(&ctrl.name)) {
                    if let Some(store) = self.data_store.get(&bs_name.to_lowercase()) {
                        ctrl.grid_columns = store.columns.clone();
                        ctrl.grid_rows = store.rows.iter().map(|row| {
                            store.columns.iter().map(|col| {
                                row.get(col).cloned().unwrap_or_default()
                            }).collect()
                        }).collect();
                        ctrl.nav_position = store.position;
                    }
                }
            } else if type_lower == "bindingnavigator" {
                if let Some(nav_info) = self.navigators.iter().find(|n| n.navigator_name.eq_ignore_ascii_case(&ctrl.name)) {
                    if let Some(store) = self.data_store.get(&nav_info.binding_source_name.to_lowercase()) {
                        ctrl.nav_position = store.position;
                        ctrl.nav_count = store.rows.len() as i32;
                    }
                }
            }
        }
    }

    /// Find the BindingSource name that a DataGridView is bound to.
    fn find_grid_binding_source(&self, grid_name: &str) -> Option<String> {
        // Check VM objects: grid.datasource should reference a BindingSource
        let vm = self.vm.borrow();
        if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
            let fo = form_obj.borrow();
            if let Some(vybe_bytecode::Value::Object(grid_obj)) = fo.properties.get(&grid_name.to_lowercase()) {
                let g = grid_obj.borrow();
                if let Some(vybe_bytecode::Value::Object(bs_ref)) = g.properties.get("datasource") {
                    let bs = bs_ref.borrow();
                    if let Some(vybe_bytecode::Value::String(name)) = bs.properties.get("__control_name") {
                        return Some(name.to_string());
                    }
                }
            }
        }
        // Fallback: check binding_sources list
        if self.binding_sources.len() == 1 {
            return Some(self.binding_sources[0].name.clone());
        }
        None
    }

    /// Navigate a BindingSource: "first", "prev", "next", "last"
    fn navigate_binding_source(&mut self, bs_name: &str, action: &str) {
        let bs_lower = bs_name.to_lowercase();
        let (new_pos, _count) = {
            let store = match self.data_store.get(&bs_lower) {
                Some(s) => s,
                None => return,
            };
            let count = store.rows.len() as i32;
            if count == 0 { return; }
            let new_pos = match action {
                "first" => 0,
                "prev" => (store.position - 1).max(0),
                "next" => (store.position + 1).min(count - 1),
                "last" => count - 1,
                _ => store.position,
            };
            (new_pos, count)
        };

        // Update position in store
        if let Some(store) = self.data_store.get_mut(&bs_lower) {
            store.position = new_pos;
        }

        // Sync bound controls
        self.sync_bound_controls(bs_name);
        self.update_data_controls();
        self.needs_redraw = true;
    }

    /// Handle click on a DataGridView row — select the clicked row.
    fn handle_grid_click(&mut self, grid_name: &str, _click_x: f64, click_y: f64) {
        let grid_ctrl = self.controls.iter().find(|c| c.name.eq_ignore_ascii_case(grid_name));
        let grid_ctrl = match grid_ctrl {
            Some(c) => c,
            None => return,
        };
        let font_size = 13.0_f32;
        let header_h = (font_size * 1.8) as i32;
        let row_h = (font_size * 1.6) as i32;
        let rel_y = (click_y - grid_ctrl.y as f64) as i32;
        if rel_y <= header_h { return; } // clicked on header
        let row_idx = ((rel_y - header_h) / row_h) as i32;

        // Find BindingSource for this grid
        if let Some(bs_name) = self.find_grid_binding_source(grid_name) {
            let bs_lower = bs_name.to_lowercase();
            let valid = self.data_store.get(&bs_lower)
                .map(|s| row_idx < s.rows.len() as i32)
                .unwrap_or(false);
            if valid {
                if let Some(store) = self.data_store.get_mut(&bs_lower) {
                    store.position = row_idx;
                }
                self.sync_bound_controls(&bs_name);
                self.update_data_controls();
                self.needs_redraw = true;
            }
        }
    }

    /// Handle click on a BindingNavigator — determine which button was clicked.
    fn handle_navigator_click(&mut self, nav_name: &str, click_x: f64, _click_y: f64) -> bool {
        let nav_ctrl = self.controls.iter().find(|c| c.name.eq_ignore_ascii_case(nav_name));
        let nav_ctrl = match nav_ctrl {
            Some(c) => c,
            None => return false,
        };
        let nav_info = self.navigators.iter().find(|n| n.navigator_name.eq_ignore_ascii_case(nav_name));
        let bs_name = match nav_info {
            Some(n) => n.binding_source_name.clone(),
            None => return false,
        };

        // Calculate which button was clicked based on relative x position
        let font_size = 13.0_f32;
        let btn_w = (font_size * 2.0) as i32;
        let gap = 2;
        let counter_w = btn_w * 2;
        let rel_x = (click_x - nav_ctrl.x as f64) as i32;

        // Button layout: |< (btn_w+gap) < (btn_w+gap) counter (counter_w+gap) > (btn_w+gap) >| (btn_w+gap) + (btn_w+gap) -
        let mut bx = 2;
        let buttons = ["first", "prev", "counter", "next", "last", "add", "remove"];
        for (i, action) in buttons.iter().enumerate() {
            let this_w = if i == 2 { counter_w } else { btn_w };
            if rel_x >= bx && rel_x < bx + this_w {
                match *action {
                    "first" | "prev" | "next" | "last" => {
                        eprintln!("[NAV] {} on {}", action, bs_name);
                        self.navigate_binding_source(&bs_name, action);
                        return true;
                    }
                    _ => return false,
                }
            }
            bx += this_w + gap;
        }
        false
    }

    #[allow(dead_code)]
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

            // Initialize data bindings — connect to DB, fill data, sync controls
            self.init_data_bindings();
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
                            grid_columns: ctrl.grid_columns.clone(),
                            grid_rows: ctrl.grid_rows.clone(),
                            nav_position: ctrl.nav_position,
                            nav_count: ctrl.nav_count,
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
                }).map(|c| (c.name.clone(), c.type_name.clone()));
                if let Some((name, type_name)) = clicked {
                    eprintln!("[SKIA-CLICK] {} ({})", name, type_name);
                    // BindingNavigator sub-button click
                    if type_name.eq_ignore_ascii_case("bindingnavigator") {
                        if self.handle_navigator_click(&name, mx, my) {
                            if let Some(window) = &self.window {
                                window.request_redraw();
                            }
                        }
                    } else {
                        // Also check if clicking on a DataGridView row
                        if type_name.eq_ignore_ascii_case("datagridview") {
                            self.handle_grid_click(&name, mx, my);
                        }
                        self.handle_click(&name);
                    }
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
    vm.register_host_fn("vybe:gui", "__dlg_show", Box::new(|_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
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
    vm.register_host_fn("vybe:gui", "msgBox", Box::new(|_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
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
    vm.register_host_fn("vybe:gui", "inputBox", Box::new(|_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
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
    _title: &str,
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
                grid_columns: Vec::new(),
                grid_rows: Vec::new(),
                nav_position: 0,
                nav_count: 0,
            }
        })
        .collect();

    // Extract data binding info from the form model
    let (data_bindings, binding_sources, navigators) = extract_binding_info(form);

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
        data_bindings,
        binding_sources,
        navigators,
        data_store: std::collections::HashMap::new(),
    };

    event_loop.run_app(&mut app).ok();
}

/// Extract data binding info from a parsed form model.
fn extract_binding_info(form: &vybe_forms::Form) -> (Vec<DataBindingEntry>, Vec<BindingSourceInfo>, Vec<NavigatorInfo>) {
    let mut data_bindings = Vec::new();
    let mut binding_sources = Vec::new();
    let mut navigators = Vec::new();

    for ctrl in &form.controls {
        let type_name = format!("{:?}", ctrl.control_type);

        // BindingSource: extract DataSource (DataAdapter name) and DataMember
        if type_name.contains("BindingSource") {
            let data_source = ctrl.properties.get_string("DataSource").unwrap_or_default().to_string();
            let data_member = ctrl.properties.get_string("DataMember").unwrap_or_default().to_string();
            if !data_source.is_empty() && !data_member.is_empty() {
                eprintln!("[BINDING] BindingSource '{}': DataSource={}, DataMember={}", ctrl.name, data_source, data_member);
                binding_sources.push(BindingSourceInfo {
                    name: ctrl.name.clone(),
                    data_adapter_name: data_source,
                    data_member,
                });
            }
        }

        // BindingNavigator: extract BindingSource reference
        if type_name.contains("BindingNavigator") {
            let bs = ctrl.properties.get_string("BindingSource").unwrap_or_default().to_string();
            if !bs.is_empty() {
                eprintln!("[BINDING] Navigator '{}' → BindingSource '{}'", ctrl.name, bs);
                navigators.push(NavigatorInfo {
                    navigator_name: ctrl.name.clone(),
                    binding_source_name: bs,
                });
            }
        }

        // DataBindings.Add on any control: extract binding entries
        let binding_source = ctrl.properties.get_string("DataBindings.Source").map(|s| s.to_string());
        if let Some(ref bs_name) = binding_source {
            if !bs_name.is_empty() {
                // Iterate properties for DataBindings.<PropName> entries
                for (key, val) in ctrl.properties.iter() {
                    let k = key.as_str();
                    if k.starts_with("DataBindings.") && k != "DataBindings.Source" {
                        let prop = &k["DataBindings.".len()..];
                        if let Some(column) = val.as_string() {
                            if !column.is_empty() {
                                eprintln!("[BINDING] {}.{} ← {}.{}", ctrl.name, prop, bs_name, column);
                                data_bindings.push(DataBindingEntry {
                                    control_name: ctrl.name.clone(),
                                    property: prop.to_string(),
                                    source_name: bs_name.clone(),
                                    column: column.to_string(),
                                });
                            }
                        }
                    }
                }
            }
        }
    }

    (data_bindings, binding_sources, navigators)
}
