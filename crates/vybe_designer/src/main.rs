#![allow(deprecated)]

use std::rc::Rc;
use tiny_skia::{Pixmap, Paint, Rect, Transform, Stroke, PathBuilder};
use winit::event::{WindowEvent, ElementState, MouseButton};
use winit::event_loop::{EventLoop, ActiveEventLoop, ControlFlow};
use winit::window::{WindowAttributes, Window};
use winit::application::ApplicationHandler;
use softbuffer::Context;

use vybe_widgets::Toolbox;
use vybe_forms::form::Form;
use vybe_forms::control::{Control, ControlType};
use cosmic_text::{FontSystem, SwashCache, Color as CosmicColor};

#[derive(Clone, Copy)]
enum ResizeHandle { TopLeft, TopRight, BottomLeft, BottomRight }

struct DesignerApp {
    window: Option<Rc<Window>>,
    context: Option<softbuffer::Context<Rc<Window>>>,
    surface: Option<softbuffer::Surface<Rc<Window>, Rc<Window>>>,
    pixmap: Option<Pixmap>,
    scale: f32,
    fs: FontSystem,
    sc: SwashCache,
    form: Form,
    toolbox: Toolbox,
    cursor_physical: (f32, f32),
    cursor_logical: (f32, f32),
    selected_idx: Option<usize>,
    dragging: bool,
    drag_offset: (f32, f32),
    resizing: Option<(usize, ResizeHandle)>,
    initial_bounds: Option<vybe_forms::control::Bounds>,
}

fn main() {
    let event_loop = EventLoop::new().expect("Failed to create event loop");
    event_loop.set_control_flow(ControlFlow::Wait);
    let mut app = DesignerApp::default();
    event_loop.run_app(&mut app).expect("Failed to run DesignerApp");
}

impl DesignerApp {
    fn request_redraw(&self) { if let Some(w) = self.window.as_ref() { w.request_redraw(); } }
}

impl Default for DesignerApp {
    fn default() -> Self {
        Self {
            window: None,
            context: None,
            surface: None,
            pixmap: None,
            scale: 1.0,
            fs: FontSystem::new(),
            sc: SwashCache::new(),
            form: Form::new("MainForm"),
            toolbox: Toolbox::new(vec![
                "Button","Label","TextBox","CheckBox","RadioButton","ComboBox","ListBox","Frame","PictureBox","RichTextBox","WebBrowser","TreeView","DataGridView","Panel","ListView","BindingNavigator","TabControl","TabPage","ProgressBar","NumericUpDown","MenuStrip","ToolStripMenuItem","ContextMenuStrip","StatusStrip","ToolStripStatusLabel","DateTimePicker","LinkLabel","ToolStrip","TrackBar","MaskedTextBox","SplitContainer","FlowLayoutPanel","TableLayoutPanel","MonthCalendar","HScrollBar","VScrollBar","ToolTip","CheckedListBox","DomainUpDown","PropertyGrid","Splitter","DataGrid","UserControl",
            ]),
            cursor_physical: (0.0,0.0),
            cursor_logical: (0.0,0.0),
            selected_idx: None,
            dragging: false,
            drag_offset: (0.0,0.0),
            resizing: None,
            initial_bounds: None,
        }
    }
}

impl ApplicationHandler for DesignerApp {
    fn resumed(&mut self, event_loop: &ActiveEventLoop) {
        let attrs = WindowAttributes::default()
            .with_title("Vybe Designer")
            .with_inner_size(winit::dpi::LogicalSize::new(1000.0, 700.0));
        let window = Rc::new(event_loop.create_window(attrs).expect("create window"));
        let context = Context::new(window.clone()).expect("softbuffer context");
        let mut surface = softbuffer::Surface::new(&context, window.clone()).expect("softbuffer surface");
        let size = window.inner_size();
        surface.resize(std::num::NonZeroU32::new(size.width).unwrap(), std::num::NonZeroU32::new(size.height).unwrap()).unwrap();
        let pix = Pixmap::new(size.width, size.height).unwrap();

        self.window = Some(window.clone());
        self.context = Some(context);
        self.surface = Some(surface);
        self.pixmap = Some(pix);
        self.scale = window.scale_factor() as f32;
    }

    fn window_event(&mut self, event_loop: &ActiveEventLoop, _id: winit::window::WindowId, event: WindowEvent) {
        let window = match self.window.as_ref() { Some(w) => w.clone(), None => return };
        match event {
            WindowEvent::CloseRequested => { event_loop.exit(); }
            WindowEvent::Resized(sz) => {
                if let Some(surface) = self.surface.as_mut() {
                    surface.resize(std::num::NonZeroU32::new(sz.width).unwrap(), std::num::NonZeroU32::new(sz.height).unwrap()).unwrap();
                }
                self.pixmap = Some(Pixmap::new(sz.width, sz.height).unwrap());
                self.request_redraw();
            }
            WindowEvent::ScaleFactorChanged { scale_factor, .. } => {
                self.scale = scale_factor as f32;
                if let Some(s) = self.window.as_ref() {
                    let si = s.inner_size();
                    self.pixmap = Some(Pixmap::new(si.width, si.height).unwrap());
                    if let Some(surf) = self.surface.as_mut() { surf.resize(std::num::NonZeroU32::new(si.width).unwrap(), std::num::NonZeroU32::new(si.height).unwrap()).unwrap(); }
                }
                self.request_redraw();
            }
            WindowEvent::CursorMoved { position, .. } => {
                self.cursor_physical = (position.x as f32, position.y as f32);
                self.cursor_logical = (position.x as f32 / self.scale, position.y as f32 / self.scale);
                // design surface origin in logical coords (must match RedrawRequested)
                let form_x_log = 200.0f32; let form_y_log = 20.0f32;
                let local_mx = self.cursor_logical.0 - form_x_log;
                let local_my = self.cursor_logical.1 - form_y_log;
                if self.dragging {
                    if let Some(idx) = self.selected_idx {
                        if let Some(ctrl) = self.form.controls.get_mut(idx) {
                            let old = ctrl.bounds.clone();
                            let nx = (local_mx - self.drag_offset.0).round() as i32;
                            let ny = (local_my - self.drag_offset.1).round() as i32;
                            ctrl.bounds.x = nx; ctrl.bounds.y = ny;
                            eprintln!("drag move ctrl {} '{}' from ({},{}) -> ({},{})", idx, ctrl.name, old.x, old.y, ctrl.bounds.x, ctrl.bounds.y);
                            self.request_redraw();
                        }
                    }
                } else if let Some((idx, handle)) = self.resizing {
                    if let Some(ctrl) = self.form.controls.get_mut(idx) {
                        if let Some(init) = self.initial_bounds {
                            let mx = local_mx; let my = local_my;
                            let x = init.x as f32; let y = init.y as f32; let w = init.width as f32; let h = init.height as f32;
                            let min_w = 12.0; let min_h = 12.0;
                            match handle {
                                ResizeHandle::TopLeft => {
                                    let old = ctrl.bounds.clone();
                                    let right = x + w; let bottom = y + h; let nx = mx.min(right - min_w); let ny = my.min(bottom - min_h);
                                    let nw = right - nx; let nh = bottom - ny;
                                    ctrl.bounds.x = nx.round() as i32; ctrl.bounds.y = ny.round() as i32; ctrl.bounds.width = nw.round() as i32; ctrl.bounds.height = nh.round() as i32;
                                    eprintln!("resize TL ctrl {} '{}' from ({},{},{},{}) -> ({},{},{},{})", idx, ctrl.name, old.x, old.y, old.width, old.height, ctrl.bounds.x, ctrl.bounds.y, ctrl.bounds.width, ctrl.bounds.height);
                                }
                                ResizeHandle::TopRight => {
                                    let old = ctrl.bounds.clone();
                                    let left = x; let bottom = y + h; let nxw = (mx - left).max(min_w); let ny = my.min(bottom - min_h); let nh = bottom - ny;
                                    ctrl.bounds.width = nxw.round() as i32; ctrl.bounds.y = ny.round() as i32; ctrl.bounds.height = nh.round() as i32;
                                    eprintln!("resize TR ctrl {} '{}' from ({},{},{},{}) -> ({},{},{},{})", idx, ctrl.name, old.x, old.y, old.width, old.height, ctrl.bounds.x, ctrl.bounds.y, ctrl.bounds.width, ctrl.bounds.height);
                                }
                                ResizeHandle::BottomLeft => {
                                    let old = ctrl.bounds.clone();
                                    let right = x + w; let nyh = (my - y).max(min_h); let nx = mx.min(right - min_w); let nw = right - nx;
                                    ctrl.bounds.x = nx.round() as i32; ctrl.bounds.width = nw.round() as i32; ctrl.bounds.height = nyh.round() as i32;
                                    eprintln!("resize BL ctrl {} '{}' from ({},{},{},{}) -> ({},{},{},{})", idx, ctrl.name, old.x, old.y, old.width, old.height, ctrl.bounds.x, ctrl.bounds.y, ctrl.bounds.width, ctrl.bounds.height);
                                }
                                ResizeHandle::BottomRight => {
                                    let old = ctrl.bounds.clone();
                                    let nxw = (mx - x).max(min_w); let nyh = (my - y).max(min_h);
                                    ctrl.bounds.width = nxw.round() as i32; ctrl.bounds.height = nyh.round() as i32;
                                    eprintln!("resize BR ctrl {} '{}' from ({},{},{},{}) -> ({},{},{},{})", idx, ctrl.name, old.x, old.y, old.width, old.height, ctrl.bounds.x, ctrl.bounds.y, ctrl.bounds.width, ctrl.bounds.height);
                                }
                            }
                            self.request_redraw();
                        }
                    }
                }
            }
            WindowEvent::MouseInput { state, button, .. } => {
                if state == ElementState::Pressed && button == MouseButton::Left {
                    if let Some((px, py)) = Some(self.cursor_physical) {
                        if let Some(idx) = self.toolbox.hit_test(px, py, 8.0 * self.scale, 8.0 * self.scale, self.scale) {
                            if let Some(item) = self.toolbox.items.get(idx) {
                                let ct = ControlType::from_name(item).unwrap_or(ControlType::Custom(item.clone()));
                                let prefix = ct.default_name_prefix();
                                let name = format!("{}{}", prefix, self.form.controls.len() + 1);
                                // store control bounds relative to form origin (not absolute)
                                let rel_x = 10 + ((self.form.controls.len() as i32) % 20) * 10;
                                let rel_y = 10 + ((self.form.controls.len() as i32) / 20) * 10;
                                let ctrl = Control::new(ct, name, rel_x, rel_y);
                                self.form.add_control(ctrl);
                                self.request_redraw();
                            }
                        } else {
                            // use local design-surface coordinates for hit-testing (logical)
                            let form_x_log = 200.0f32; let form_y_log = 20.0f32;
                            let mx = self.cursor_logical.0 - form_x_log; let my = self.cursor_logical.1 - form_y_log;
                            eprintln!("click: physical=({:.1},{:.1}) logical=({:.2},{:.2})", self.cursor_physical.0, self.cursor_physical.1, self.cursor_logical.0, self.cursor_logical.1);
                            // dump all control positions (logical and physical) to help debug mismatch
                            let form_x_log = 200.0f32; let form_y_log = 20.0f32;
                            for (ci, cc) in self.form.controls.iter().enumerate() {
                                let cb = cc.bounds;
                                let clx = form_x_log + cb.x as f32;
                                let cly = form_y_log + cb.y as f32;
                                let cpx = clx * self.scale;
                                let cpy = cly * self.scale;
                                eprintln!("ctrl {} '{}' logical=({:.1},{:.1}) size=({}x{}) phys=({:.1},{:.1})", ci, cc.name, clx, cly, cb.width, cb.height, cpx, cpy);
                            }
                            let mut found = None;
                            for (i, c) in self.form.controls.iter().enumerate() {
                                let b = c.bounds;
                                if mx >= b.x as f32 && mx <= (b.x + b.width) as f32 && my >= b.y as f32 && my <= (b.y + b.height) as f32 { found = Some(i); break; }
                            }
                            if let Some(idx) = found {
                                let bound = self.form.controls[idx].bounds;
                                eprintln!("hit control {} bounds=({},{} {} {})", idx, bound.x, bound.y, bound.width, bound.height);
                                // also print hit-test rect in physical pixels
                                let form_x_log = 200.0f32; let form_y_log = 20.0f32;
                                let hit_px = (form_x_log + bound.x as f32) * self.scale;
                                let hit_py = (form_y_log + bound.y as f32) * self.scale;
                                let hit_pw = (bound.width as f32) * self.scale;
                                let hit_ph = (bound.height as f32) * self.scale;
                                eprintln!("hit rect phys=({:.1},{:.1}) size=({:.1}x{:.1})", hit_px, hit_py, hit_pw, hit_ph);
                                let b = self.form.controls[idx].bounds; let x = b.x as f32; let y = b.y as f32; let w = b.width as f32; let h = b.height as f32;
                                let hs = 6.0; let half = hs / 2.0;
                                let corners = [ (x - half, y - half, ResizeHandle::TopLeft), (x + w - half, y - half, ResizeHandle::TopRight), (x - half, y + h - half, ResizeHandle::BottomLeft), (x + w - half, y + h - half, ResizeHandle::BottomRight) ];
                                let mut handle_hit: Option<ResizeHandle> = None;
                                for (hx, hy, hdl) in &corners { if mx >= *hx && mx <= (*hx + hs) && my >= *hy && my <= (*hy + hs) { handle_hit = Some(*hdl); break; } }
                                if let Some(h) = handle_hit { self.selected_idx = Some(idx); self.resizing = Some((idx, h)); self.initial_bounds = Some(self.form.controls[idx].bounds.clone()); self.dragging = false; self.request_redraw(); }
                                else { self.selected_idx = Some(idx); self.dragging = true; let b = self.form.controls[idx].bounds; self.drag_offset = (mx - b.x as f32, my - b.y as f32); self.request_redraw(); }
                            } else { self.selected_idx = None; self.dragging = false; }
                        }
                    }
                } else if state == ElementState::Released && button == MouseButton::Left { self.dragging = false; self.resizing = None; self.initial_bounds = None; }
            }
            WindowEvent::RedrawRequested => {
                // perform rendering
                if let (Some(pixmap), Some(surface)) = (self.pixmap.as_mut(), self.surface.as_mut()) {
                    pixmap.fill(tiny_skia::Color::from_rgba8(30,30,30,255));
                    // Use logical coordinates for widget/text APIs and physical (px) for direct tiny-skia drawing
                    let logical_canvas_w = (pixmap.width() as f32) / self.scale;
                    let logical_canvas_h = (pixmap.height() as f32) / self.scale;

                    // toolbox: draw background in physical pixels, draw text in logical coords
                    let toolbox_x_log = 8.0; let toolbox_y_log = 8.0;
                    let toolbox_x = toolbox_x_log * self.scale; let toolbox_y = toolbox_y_log * self.scale;
                    self.toolbox.paint(pixmap, toolbox_x, toolbox_y, self.scale);
                    // toolbox labels (logical coords)
                    let mut iy_log = toolbox_y_log + 8.0; let label_size_log = 14.0; let text_col = CosmicColor::rgb(240,240,240);
                    for item in &self.toolbox.items { let item_h_log = 28.0; let it_y_log = iy_log + (item_h_log - (label_size_log * 1.2)) / 2.0; let text_x_px = (toolbox_x_log + 12.0) * self.scale; let text_y_px = it_y_log * self.scale; vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, item, text_x_px, text_y_px, text_col, self.scale); iy_log += item_h_log; }
                    // right-side list (logical coords)
                    let right_w_log = 240.0; let right_x_log = logical_canvas_w - right_w_log + 8.0;
                    vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, "Controls", right_x_log * self.scale, (toolbox_y_log + 6.0) * self.scale, CosmicColor::rgb(240,240,240), self.scale);
                    let mut ry_log = toolbox_y_log + 36.0; let item_h_log = 22.0;
                    for c in &self.form.controls { let label = format!("{} ({})", c.name, c.control_type.as_str()); let label_x_px = (right_x_log + 4.0) * self.scale; let label_y_px = ry_log * self.scale; vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, &label, label_x_px, label_y_px, CosmicColor::rgb(230,230,230), self.scale); ry_log += item_h_log + 6.0; }
                    // design surface placeholder (physical coords)
                    let mut pb = PathBuilder::new(); let d_form_x_log = 200.0; let d_form_y_log = 20.0; let d_form_w_log = logical_canvas_w - 240.0; let d_form_h_log = logical_canvas_h - 60.0;
                    let d_form_x = d_form_x_log * self.scale; let d_form_y = d_form_y_log * self.scale; let d_form_w = d_form_w_log * self.scale; let d_form_h = d_form_h_log * self.scale;
                    pb.push_rect(Rect::from_xywh(d_form_x, d_form_y, d_form_w, d_form_h).unwrap()); if let Some(path) = pb.finish() { let mut p = Paint::default(); p.set_color_rgba8(200,200,200,40); pixmap.stroke_path(&path, &p, &Stroke::default(), Transform::identity(), None); }

                    // render controls (wired properties for lists/selects/datagrid)
                    // Compute logical (lx,ly,lw,lh) and physical (px,py,pw,ph) coords.
                    let form_x_log = d_form_x_log; let form_y_log = d_form_y_log;
                    for (i, ctrl) in self.form.controls.iter().enumerate() {
                        let b = ctrl.bounds;
                        let lx = form_x_log + b.x as f32;
                        let ly = form_y_log + b.y as f32;
                        let lw = b.width as f32;
                        let lh = b.height as f32;
                        let x = lx * self.scale;
                        let y = ly * self.scale;
                        let w = lw * self.scale;
                        let h = lh * self.scale;
                        match &ctrl.control_type {
                            ControlType::Button => {
                                let label = ctrl.get_text().unwrap_or(&ctrl.name);
                                let mut btn = vybe_widgets::Button::new(label);
                                btn.width = lw; btn.height = lh; btn.focused = Some(i) == self.selected_idx; btn.paint(pixmap, lx, ly, self.scale);
                                // Draw the button's inner text (logical coords)
                                let tx = lx + 6.0;
                                let ty = ly + (lh - 14.0) / 2.0 + 2.0;
                                vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, label, tx * self.scale, ty * self.scale, CosmicColor::rgb(20,20,20), self.scale);
                                // Draw the control name above the control (logical coords)
                                let name_ty = ly - 14.0 - 4.0;
                                vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, &ctrl.name, lx * self.scale, name_ty * self.scale, CosmicColor::rgb(240,240,240), self.scale);
                            }
                            ControlType::TextBox => { let mut tf = vybe_widgets::TextInput::new(); tf.width = lw; tf.height = lh; tf.focused = Some(i) == self.selected_idx; tf.paint_border(pixmap, lx, ly, self.scale); let label = ctrl.get_text().unwrap_or(&ctrl.name); let tx = lx + 6.0; let ty = ly + (lh - 14.0) / 2.0 + 2.0; vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, label, tx * self.scale, ty * self.scale, CosmicColor::rgb(20,20,20), self.scale); }
                            ControlType::Label => { let label = ctrl.get_text().unwrap_or(&ctrl.name); let tx = lx + 4.0; let ty = ly + (lh - 14.0) / 2.0 + 2.0; vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, label, tx * self.scale, ty * self.scale, CosmicColor::rgb(20,20,20), self.scale); }
                            ControlType::CheckBox => { let mut cb = vybe_widgets::Checkbox::new(""); cb.size = lh.min(lw).min(16.0); cb.checked = ctrl.properties.get_bool("Checked").unwrap_or(false); cb.focused = Some(i) == self.selected_idx; cb.paint(pixmap, lx, ly, self.scale); let label = ctrl.get_text().unwrap_or(&ctrl.name); let tx = lx + cb.size + 6.0; let ty = ly + (lh - 14.0) / 2.0 + 2.0; vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, label, tx * self.scale, ty * self.scale, CosmicColor::rgb(20,20,20), self.scale); }
                            ControlType::TrackBar => { let mut sl = vybe_widgets::Slider::new(0.0, 100.0, 50.0); sl.width = lw; sl.height = lh; sl.paint(pixmap, lx, ly, self.scale); }
                            ControlType::RadioButton => { let mut r = vybe_widgets::Radio::new(""); r.focused = Some(i) == self.selected_idx; r.selected = ctrl.properties.get_bool("Checked").unwrap_or(false); r.paint(pixmap, lx, ly, self.scale); let label = ctrl.get_text().unwrap_or(&ctrl.name); let tx = lx + r.size + 6.0; let ty = ly + (lh - 14.0) / 2.0 + 2.0; vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, label, tx * self.scale, ty * self.scale, CosmicColor::rgb(20,20,20), self.scale); }
                            ControlType::Panel => { let mut p = vybe_widgets::Panel::new(); p.width = lw; p.height = lh; p.paint(pixmap, lx, ly, self.scale); }
                            ControlType::ProgressBar => { let mut pbw = vybe_widgets::ProgressBar::new(); pbw.width = lw; pbw.height = lh; if let Some(v) = ctrl.properties.get_int("Value") { pbw.value = (v as f32) / 100.0; } pbw.paint(pixmap, lx, ly, self.scale); }
                            ControlType::ListBox => { let mut lb = vybe_widgets::ListBox::new(); if let Some(items) = ctrl.properties.get_string_array("Items") { lb.items = items.clone(); } lb.width = lw; lb.height = lh; lb.paint(pixmap, lx, ly, self.scale); }
                            ControlType::PictureBox => { let mut pic = vybe_widgets::PictureBox::new(); pic.width = lw; pic.height = lh; pic.paint(pixmap, lx, ly, self.scale); }
                            ControlType::NumericUpDown => { let mut nud = vybe_widgets::NumericUpDown::new(); nud.width = lw; nud.height = lh; nud.paint(pixmap, lx, ly, self.scale); }
                            ControlType::LinkLabel => { let label = ctrl.get_text().unwrap_or(&ctrl.name); let mut ll = vybe_widgets::LinkLabel::new(label); ll.width = lw; ll.height = lh; ll.paint(pixmap, lx, ly, self.scale); }
                            ControlType::TreeView => { let mut tv = vybe_widgets::TreeView::new(".", self.scale); tv.render(pixmap, &mut self.fs, &mut self.sc, lx, ly, lw, CosmicColor::rgb(200,200,200), (60,60,60,120)); }
                            ControlType::ComboBox => { let options = ctrl.properties.get_string_array("Items").map(|v| v.clone()).unwrap_or_else(|| vec![]); let mut sel = vybe_widgets::Select::new(options); if let Some(idx) = ctrl.properties.get_int("SelectedIndex") { sel.selected_index = idx.max(0) as usize; } sel.width = lw; sel.height = lh; sel.paint(pixmap, lx, ly, self.scale); let txt = sel.selected_text().to_string(); vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, &txt, (lx + 6.0) * self.scale, (ly + (lh - 14.0)/2.0 + 2.0) * self.scale, CosmicColor::rgb(20,20,20), self.scale); }
                            ControlType::Frame => { let title = ctrl.get_text().map(|s| s.to_string()).unwrap_or_else(|| ctrl.name.clone()); let mut gb = vybe_widgets::GroupBox::new(title); gb.width = lw; gb.height = lh; gb.paint(pixmap, lx, ly, self.scale); }
                            ControlType::RichTextBox => { let mut p = vybe_widgets::Panel::new(); p.width = lw; p.height = lh; p.paint(pixmap, lx, ly, self.scale); }
                            ControlType::WebBrowser => { let mut pic = vybe_widgets::PictureBox::new(); pic.width = lw; pic.height = lh; pic.paint(pixmap, lx, ly, self.scale); }
                            ControlType::DataGridView | ControlType::DataGrid => { let cols_vec = ctrl.properties.get_string_array("Columns").map(|v| v.iter().map(|s| s.as_str()).collect::<Vec<&str>>()).unwrap_or_default(); let mut dg = if cols_vec.is_empty() { vybe_widgets::DataGrid::new(&[]) } else { vybe_widgets::DataGrid::new(&cols_vec) }; dg.width = lw; dg.height = lh; dg.paint(pixmap, lx, ly, self.scale); }
                            ControlType::ListView => { let mut lv = vybe_widgets::ListView::new(); lv.width = lw; lv.height = lh; lv.paint(pixmap, lx, ly, self.scale); }
                            ControlType::BindingNavigator => { let mut ts = vybe_widgets::ToolStrip::new(); ts.width = lw; ts.height = lh; ts.paint(pixmap, lx, ly, self.scale); }
                            ControlType::TabControl => { let mut tabs = vybe_widgets::Tabs::new(&["Tab1","Tab2"]); tabs.width = lw; tabs.height = lh; tabs.paint(pixmap, lx, ly, self.scale); }
                            ControlType::TabPage => { let mut p = vybe_widgets::Panel::new(); p.width = lw; p.height = lh; p.paint(pixmap, lx, ly, self.scale); }
                            ControlType::MenuStrip => { let mut ms = vybe_widgets::MenuStrip::new(); ms.width = lw; ms.height = lh; ms.paint(pixmap, lx, ly, self.scale); }
                            ControlType::ToolStrip | ControlType::ToolStripButton | ControlType::ToolStripLabel => { let mut ts = vybe_widgets::ToolStrip::new(); ts.width = lw; ts.height = lh; ts.paint(pixmap, lx, ly, self.scale); }
                            ControlType::StatusStrip => { let mut ss = vybe_widgets::StatusStrip::new(); ss.width = lw; ss.height = lh; ss.paint(pixmap, lx, ly, self.scale); }
                            ControlType::DateTimePicker => { let mut dt = vybe_widgets::DateTimePicker::new(); dt.width = lw; dt.height = lh; dt.paint(pixmap, lx, ly, self.scale); }
                            ControlType::MaskedTextBox => { let mut mt = vybe_widgets::MaskedTextBox::new(); mt.width = lw; mt.height = lh; mt.paint(pixmap, lx, ly, self.scale); }
                            ControlType::SplitContainer => { let horiz = ctrl.properties.get_bool("Horizontal").unwrap_or(true); let mut scn = vybe_widgets::SplitContainer::new(horiz); scn.width = lw; scn.height = lh; scn.paint(pixmap, lx, ly, self.scale); }
                            ControlType::FlowLayoutPanel => { let mut fl = vybe_widgets::FlowLayoutPanel::new(); fl.width = lw; fl.height = lh; fl.paint(pixmap, lx, ly, self.scale); }
                            ControlType::TableLayoutPanel => { let cols = ctrl.properties.get_int("Cols").unwrap_or(2) as usize; let rows = ctrl.properties.get_int("Rows").unwrap_or(2) as usize; let mut tl = vybe_widgets::TableLayoutPanel::new(cols, rows); tl.width = lw; tl.height = lh; tl.paint(pixmap, lx, ly, self.scale); }
                            ControlType::MonthCalendar => { let mut mc = vybe_widgets::MonthCalendar::new(); mc.width = lw; mc.height = lh; mc.paint(pixmap, lx, ly, self.scale); }
                            ControlType::HScrollBar => { let mut sb = vybe_widgets::ScrollBar::new(false); sb.width = lw; sb.height = lh; sb.paint(pixmap, lx, ly, self.scale); }
                            ControlType::VScrollBar => { let mut sb = vybe_widgets::ScrollBar::new(true); sb.width = lw; sb.height = lh; sb.paint(pixmap, lx, ly, self.scale); }
                            ControlType::CheckedListBox => { let mut lb = vybe_widgets::ListBox::new(); if let Some(items) = ctrl.properties.get_string_array("Items") { lb.items = items.clone(); } lb.width = lw; lb.height = lh; lb.paint(pixmap, lx, ly, self.scale); }
                            ControlType::ToolTip => { let label = ctrl.get_text().unwrap_or(&ctrl.name); let mut pb = PathBuilder::new(); pb.push_rect(Rect::from_xywh(x, y - 20.0 * self.scale, w, 18.0 * self.scale).unwrap()); if let Some(pth) = pb.finish() { let mut bg = Paint::default(); bg.set_color_rgba8(40,40,40,220); pixmap.fill_path(&pth, &bg, tiny_skia::FillRule::Winding, Transform::identity(), None); vybe_widgets::tree_view::TreeView::draw_text_static_internal(pixmap, &mut self.fs, &mut self.sc, label, x + 4.0 * self.scale, y - 16.0 * self.scale, CosmicColor::rgb(230,230,230), self.scale); } }
                            ControlType::DomainUpDown => { let mut dud = vybe_widgets::NumericUpDown::new(); dud.width = lw; dud.height = lh; dud.paint(pixmap, lx, ly, self.scale); }
                            ControlType::PropertyGrid => { let mut pg = vybe_widgets::Panel::new(); pg.width = lw; pg.height = lh; pg.paint(pixmap, lx, ly, self.scale); }
                            ControlType::Splitter => { let horiz = ctrl.properties.get_bool("Horizontal").unwrap_or(true); let mut scn = vybe_widgets::SplitContainer::new(horiz); scn.width = lw; scn.height = lh; scn.paint(pixmap, lx, ly, self.scale); }
                            ControlType::UserControl => { let mut p = vybe_widgets::Panel::new(); p.width = lw; p.height = lh; p.paint(pixmap, lx, ly, self.scale); }
                            ControlType::ToolStripSeparator | ControlType::ToolStripComboBox | ControlType::ToolStripDropDownButton | ControlType::ToolStripSplitButton | ControlType::ToolStripTextBox | ControlType::ToolStripProgressBar => { let mut tsu = vybe_widgets::ToolStrip::new(); tsu.width = lw; tsu.height = lh; tsu.paint(pixmap, lx, ly, self.scale); }
                            _ => {
                                let mut pb2 = PathBuilder::new(); pb2.push_rect(Rect::from_xywh(x, y, w, h).unwrap()); if let Some(pth) = pb2.finish() { let mut fill = Paint::default(); fill.set_color_rgba8(240,240,255,255); pixmap.fill_path(&pth, &fill, tiny_skia::FillRule::Winding, Transform::identity(), None); let mut stroke = Paint::default(); if Some(i) == self.selected_idx { stroke.set_color_rgba8(255, 100, 40, 220); } else { stroke.set_color_rgba8(60,60,60,200); } pixmap.stroke_path(&pth, &stroke, &Stroke::default(), Transform::identity(), None); }
                            }
                        }
                    }

                    // Draw selection handles (convert logical bounds to physical)
                    if let Some(idx) = self.selected_idx {
                        if let Some(ctrl) = self.form.controls.get(idx) {
                            let b = ctrl.bounds;
                            let x = (form_x_log + b.x as f32) * self.scale;
                            let y = (form_y_log + b.y as f32) * self.scale;
                            let w = (b.width as f32) * self.scale;
                            let h = (b.height as f32) * self.scale;
                            let hs = 6.0 * self.scale;
                            let corners = [ (x - hs/2.0, y - hs/2.0), (x + w - hs/2.0, y - hs/2.0), (x - hs/2.0, y + h - hs/2.0), (x + w - hs/2.0, y + h - hs/2.0) ];
                            for (hx, hy) in &corners { let mut pbh = PathBuilder::new(); pbh.push_rect(Rect::from_xywh(*hx, *hy, hs, hs).unwrap()); if let Some(pth) = pbh.finish() { let mut fill = Paint::default(); fill.set_color_rgba8(255,100,40,220); pixmap.fill_path(&pth, &fill, tiny_skia::FillRule::Winding, Transform::identity(), None); let mut stroke = Paint::default(); stroke.set_color_rgba8(0,0,0,120); pixmap.stroke_path(&pth, &stroke, &Stroke::default(), Transform::identity(), None); } }
                        }
                    }

                    // Debug overlays: draw hit-test rects (magenta) and paint rects (cyan) to visualize mismatch
                    for (i, ctrl) in self.form.controls.iter().enumerate() {
                        let b = ctrl.bounds;
                        let hit_x = (form_x_log + b.x as f32) * self.scale;
                        let hit_y = (form_y_log + b.y as f32) * self.scale;
                        let hit_w = (b.width as f32) * self.scale;
                        let hit_h = (b.height as f32) * self.scale;
                        // magenta = hit-test rect
                        let mut pb_hit = PathBuilder::new(); if let Some(r) = tiny_skia::Rect::from_xywh(hit_x, hit_y, hit_w, hit_h) { pb_hit.push_rect(r); }
                        if let Some(pth) = pb_hit.finish() { let mut stroke = Paint::default(); stroke.set_color_rgba8(255, 0, 255, 160); pixmap.stroke_path(&pth, &stroke, &Stroke { width: 1.0, ..Default::default() }, Transform::identity(), None); }
                        // cyan = paint rect (should be same)
                        let mut pb_p = PathBuilder::new(); if let Some(r2) = tiny_skia::Rect::from_xywh(hit_x, hit_y, hit_w, hit_h) { pb_p.push_rect(r2); }
                        if let Some(pth2) = pb_p.finish() { let mut stroke2 = Paint::default(); stroke2.set_color_rgba8(0, 255, 255, 120); pixmap.stroke_path(&pth2, &stroke2, &Stroke { width: 1.0, ..Default::default() }, Transform::identity(), None); }

                    }

                    // Blit to window buffer
                    if let Some(buf_surf) = self.surface.as_mut() {
                        let mut buffer = buf_surf.buffer_mut().unwrap(); for (i, px) in pixmap.pixels().iter().enumerate() { buffer[i] = (px.alpha() as u32) << 24 | (px.red() as u32) << 16 | (px.green() as u32) << 8 | (px.blue() as u32); } buffer.present().unwrap();
                    }
                }
            }
            _ => {}
        }
    }
}
