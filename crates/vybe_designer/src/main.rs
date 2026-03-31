use std::rc::Rc;
use tiny_skia::{Pixmap, Paint, Rect, Transform, Stroke, PathBuilder, PixmapPaint, ColorU8};
use winit::event::{Event, WindowEvent, ElementState, MouseButton};
use winit::event_loop::EventLoop;
use winit::window::WindowAttributes;
use softbuffer::Context;

use vybe_widgets::Toolbox;
use vybe_forms::form::Form;
use vybe_forms::control::{Control, ControlType};
use cosmic_text::{FontSystem, SwashCache, Color as CosmicColor};

fn main() {
    let event_loop = EventLoop::new().unwrap();
    let attrs = WindowAttributes::default()
        .with_title("Vybe Designer")
        .with_inner_size(winit::dpi::LogicalSize::new(1000.0, 700.0));
    let window = Rc::new(event_loop.create_window(attrs).unwrap());

    let context = Context::new(window.clone()).expect("softbuffer context");
    let mut surface = softbuffer::Surface::new(&context, window.clone()).expect("softbuffer surface");
    let mut size = window.inner_size();
    surface.resize(std::num::NonZeroU32::new(size.width).unwrap(), std::num::NonZeroU32::new(size.height).unwrap()).unwrap();

    let mut scale = window.scale_factor() as f32;
    let mut pixmap = Pixmap::new(size.width, size.height).unwrap();

    // Full visual control list for the toolbox
    let mut toolbox = Toolbox::new(vec![
        "Button","Label","TextBox","CheckBox","RadioButton","ComboBox","ListBox","Frame","PictureBox","RichTextBox","WebBrowser","TreeView","DataGridView","Panel","ListView","BindingNavigator","TabControl","TabPage","ProgressBar","NumericUpDown","MenuStrip","ToolStripMenuItem","ContextMenuStrip","StatusStrip","ToolStripStatusLabel","DateTimePicker","LinkLabel","ToolStrip","TrackBar","MaskedTextBox","SplitContainer","FlowLayoutPanel","TableLayoutPanel","MonthCalendar","HScrollBar","VScrollBar","ToolTip","CheckedListBox","DomainUpDown","PropertyGrid","Splitter","DataGrid","UserControl",
    ]);

    // cosmic-text objects for text shaping/rendering
    let mut fs = FontSystem::new();
    let mut sc = SwashCache::new();

    let mut needs_redraw = true;
    let mut form = Form::new("MainForm");
    // mouse state: physical (pixels) and logical (DIP)
    let mut cursor_physical = (0.0f32, 0.0f32);
    let mut cursor_logical = (0.0f32, 0.0f32);
    let mut selected_idx: Option<usize> = None;
    let mut dragging = false;
    let mut drag_offset = (0.0f32, 0.0f32);
    // Resizing state: (control index, handle)
    #[derive(Clone, Copy)]
    enum ResizeHandle { TopLeft, TopRight, BottomLeft, BottomRight }
    let mut resizing: Option<(usize, ResizeHandle)> = None;
    let mut initial_bounds: Option<vybe_forms::control::Bounds> = None;

    let _ = event_loop.run(move |event, event_loop| {
        match event {
            Event::WindowEvent { event, .. } => match event {
                WindowEvent::CloseRequested => {
                    event_loop.exit();
                }
                WindowEvent::Resized(sz) => {
                    size = sz;
                    surface.resize(std::num::NonZeroU32::new(sz.width).unwrap(), std::num::NonZeroU32::new(sz.height).unwrap()).unwrap();
                    pixmap = Pixmap::new(sz.width, sz.height).unwrap();
                    window.request_redraw();
                }
                WindowEvent::ScaleFactorChanged { scale_factor, .. } => {
                    scale = scale_factor as f32;
                    // Recreate pixmap at new physical size
                    let s = window.inner_size();
                    pixmap = Pixmap::new(s.width, s.height).unwrap();
                    surface.resize(std::num::NonZeroU32::new(s.width).unwrap(), std::num::NonZeroU32::new(s.height).unwrap()).unwrap();
                    window.request_redraw();
                }
                WindowEvent::CursorMoved { position, .. } => {
                    cursor_physical = (position.x as f32, position.y as f32);
                    cursor_logical = (position.x as f32 / scale, position.y as f32 / scale);
                    if dragging {
                        if let Some(idx) = selected_idx {
                            if let Some(ctrl) = form.controls.get_mut(idx) {
                                let nx = (cursor_logical.0 - drag_offset.0).round() as i32;
                                let ny = (cursor_logical.1 - drag_offset.1).round() as i32;
                                ctrl.bounds.x = nx;
                                ctrl.bounds.y = ny;
                                window.request_redraw();
                            }
                        }
                    } else if let Some((idx, handle)) = resizing {
                        // Resize the control based on handle and current logical cursor
                        if let Some(ctrl) = form.controls.get_mut(idx) {
                            if let Some(init) = initial_bounds {
                                let mx = cursor_logical.0;
                                let my = cursor_logical.1;
                                let mut x = init.x as f32;
                                let mut y = init.y as f32;
                                let mut w = init.width as f32;
                                let mut h = init.height as f32;
                                let min_w = 12.0; let min_h = 12.0;
                                match handle {
                                    ResizeHandle::TopLeft => {
                                        let right = x + w;
                                        let bottom = y + h;
                                        let nx = mx.min(right - min_w);
                                        let ny = my.min(bottom - min_h);
                                        let nw = right - nx;
                                        let nh = bottom - ny;
                                        ctrl.bounds.x = nx.round() as i32;
                                        ctrl.bounds.y = ny.round() as i32;
                                        ctrl.bounds.width = nw.round() as i32;
                                        ctrl.bounds.height = nh.round() as i32;
                                    }
                                    ResizeHandle::TopRight => {
                                        let left = x;
                                        let bottom = y + h;
                                        let nxw = (mx - left).max(min_w);
                                        let ny = my.min(bottom - min_h);
                                        let nh = bottom - ny;
                                        ctrl.bounds.width = nxw.round() as i32;
                                        ctrl.bounds.y = ny.round() as i32;
                                        ctrl.bounds.height = nh.round() as i32;
                                    }
                                    ResizeHandle::BottomLeft => {
                                        let right = x + w;
                                        let nyh = (my - y).max(min_h);
                                        let nx = mx.min(right - min_w);
                                        let nw = right - nx;
                                        ctrl.bounds.x = nx.round() as i32;
                                        ctrl.bounds.width = nw.round() as i32;
                                        ctrl.bounds.height = nyh.round() as i32;
                                    }
                                    ResizeHandle::BottomRight => {
                                        let nxw = (mx - x).max(min_w);
                                        let nyh = (my - y).max(min_h);
                                        ctrl.bounds.width = nxw.round() as i32;
                                        ctrl.bounds.height = nyh.round() as i32;
                                    }
                                }
                                window.request_redraw();
                            }
                        }
                    }
                }
                WindowEvent::MouseInput { state, button, .. } => {
                    if state == ElementState::Pressed && button == MouseButton::Left {
                        // Check toolbox first (uses physical coords)
                        if let Some(idx) = toolbox.hit_test(cursor_physical.0, cursor_physical.1, 8.0 * scale, 8.0 * scale, scale) {
                            if let Some(item) = toolbox.items.get(idx) {
                                let ct = ControlType::from_name(item).unwrap_or(ControlType::Custom(item.clone()));
                                let prefix = ct.default_name_prefix();
                                let name = format!("{}{}", prefix, form.controls.len() + 1);
                                // Place at default logical position
                                let logical_form_x = 200.0f32;
                                let logical_form_y = 20.0f32;
                                let x = (logical_form_x as i32) + 10 + ((form.controls.len() as i32) % 20) * 10;
                                let y = (logical_form_y as i32) + 10 + ((form.controls.len() as i32) / 20) * 10;
                                let ctrl = Control::new(ct, name, x, y);
                                form.add_control(ctrl);
                                window.request_redraw();
                            }
                        } else {
                            // Click on existing controls? use logical coords for hit-testing
                            let mx = cursor_logical.0;
                            let my = cursor_logical.1;
                            let mut found = None;
                            for (i, c) in form.controls.iter().enumerate() {
                                let b = c.bounds;
                                if mx >= b.x as f32 && mx <= (b.x + b.width) as f32 && my >= b.y as f32 && my <= (b.y + b.height) as f32 {
                                    found = Some(i);
                                    break;
                                }
                            }
                            if let Some(idx) = found {
                                // Check for handle hit first
                                let b = form.controls[idx].bounds;
                                let x = b.x as f32; let y = b.y as f32; let w = b.width as f32; let h = b.height as f32;
                                let hs = 6.0; let half = hs / 2.0;
                                let corners = [
                                    (x - half, y - half, ResizeHandle::TopLeft),
                                    (x + w - half, y - half, ResizeHandle::TopRight),
                                    (x - half, y + h - half, ResizeHandle::BottomLeft),
                                    (x + w - half, y + h - half, ResizeHandle::BottomRight),
                                ];
                                let mut handle_hit: Option<ResizeHandle> = None;
                                for (hx, hy, hdl) in &corners {
                                    if mx >= *hx && mx <= (*hx + hs) && my >= *hy && my <= (*hy + hs) {
                                        handle_hit = Some(*hdl);
                                        break;
                                    }
                                }
                                if let Some(h) = handle_hit {
                                    selected_idx = Some(idx);
                                    resizing = Some((idx, h));
                                    initial_bounds = Some(form.controls[idx].bounds.clone());
                                    dragging = false;
                                    window.request_redraw();
                                } else {
                                    selected_idx = Some(idx);
                                    dragging = true;
                                    let b = form.controls[idx].bounds;
                                    drag_offset = (mx - b.x as f32, my - b.y as f32);
                                    window.request_redraw();
                                }
                            } else {
                                selected_idx = None;
                                dragging = false;
                            }
                        }
                    } else if state == ElementState::Released && button == MouseButton::Left {
                        dragging = false;
                        resizing = None;
                        initial_bounds = None;
                    }
                }
                WindowEvent::RedrawRequested => {
                    // Clear
                    pixmap.fill(tiny_skia::Color::from_rgba8(30,30,30,255));

                    // Draw toolbox at left (scaled)
                    let toolbox_x = 8.0 * scale;
                    let toolbox_y = 8.0 * scale;
                    toolbox.paint(&mut pixmap, toolbox_x, toolbox_y, scale);
                    // Draw toolbox labels using cosmic-text (centered vertically in item rect)
                    let mut iy = toolbox_y + 8.0 * scale;
                    let label_size = 14.0 * scale;
                    let text_col = CosmicColor::rgb(240,240,240);
                    for item in &toolbox.items {
                        let item_h = 28.0 * scale;
                        let it_y = iy + (item_h - (label_size * 1.2)) / 2.0; // vertical center
                        vybe_widgets::tree_view::TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, item, toolbox_x + 12.0 * scale, it_y, text_col, scale);
                        iy += item_h;
                    }

                    // Right-side control list
                    let right_w = 240.0 * scale;
                    let form_x = toolbox_x + 160.0 * scale + 16.0 * scale; // left margin + toolbox width + gap
                    let form_y = 20.0 * scale;
                    let _form_w = (pixmap.width() as f32) - form_x - right_w - 16.0 * scale;
                    let right_x = (pixmap.width() as f32) - right_w + 8.0 * scale;
                    // Draw splitter
                    let splitter_x = (pixmap.width() as f32) - right_w - 8.0 * scale;
                    let mut pb_s = PathBuilder::new(); pb_s.push_rect(tiny_skia::Rect::from_xywh(splitter_x, 0.0, 6.0 * scale, pixmap.height() as f32).unwrap());
                    if let Some(pth) = pb_s.finish() { let mut p = Paint::default(); p.set_color_rgba8(80,80,80,200); pixmap.fill_path(&pth, &p, tiny_skia::FillRule::Winding, Transform::identity(), None); }

                    // Draw control list panel background
                    let mut pb_r = PathBuilder::new(); pb_r.push_rect(tiny_skia::Rect::from_xywh(right_x - 8.0*scale, toolbox_y, right_w - 8.0*scale, (pixmap.height() as f32) - 16.0*scale).unwrap());
                    if let Some(pth) = pb_r.finish() { let mut p = Paint::default(); p.set_color_rgba8(50,50,50,220); pixmap.fill_path(&pth, &p, tiny_skia::FillRule::Winding, Transform::identity(), None); }
                    // List header
                    vybe_widgets::tree_view::TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, "Controls", right_x, toolbox_y + 6.0*scale, CosmicColor::rgb(240,240,240), scale);
                    // Render control list items
                    let mut ry = toolbox_y + 36.0 * scale;
                    let item_h = 22.0 * scale;
                    for c in &form.controls {
                        let label = format!("{} ({})", c.name, c.control_type.as_str());
                        vybe_widgets::tree_view::TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, &label, right_x + 4.0*scale, ry + 0.0*scale, CosmicColor::rgb(230,230,230), scale);
                        ry += item_h + 6.0 * scale;
                    }

                    // Draw design surface placeholder (logical coords scaled)
                    let mut pb = PathBuilder::new();
                    let d_form_x = 200.0 * scale; let d_form_y = 20.0 * scale; let d_form_w = (pixmap.width() as f32) - 240.0 * scale; let d_form_h = (pixmap.height() as f32) - 60.0 * scale;
                    pb.push_rect(Rect::from_xywh(d_form_x, d_form_y, d_form_w, d_form_h).unwrap());
                    if let Some(path) = pb.finish() {
                        let mut p = Paint::default(); p.set_color_rgba8(200,200,200,40);
                        pixmap.stroke_path(&path, &p, &Stroke::default(), Transform::identity(), None);
                    }

                    // Render form controls using `vybe_widgets` paint functions (logical coords)
                    for (i, ctrl) in form.controls.iter().enumerate() {
                        let b = ctrl.bounds;
                        let x = b.x as f32;
                        let y = b.y as f32;
                        let w = b.width as f32;
                        let h = b.height as f32;

                        match &ctrl.control_type {
                            ControlType::Button => {
                                let label = ctrl.get_text().unwrap_or(&ctrl.name);
                                let mut btn = vybe_widgets::Button::new(label);
                                btn.width = w;
                                btn.height = h;
                                btn.focused = Some(i) == selected_idx;
                                btn.paint(&mut pixmap, x, y, scale);
                                // Draw label (caller handles text layout)
                                let tx = x + 6.0;
                                let ty = y + (h - 14.0) / 2.0 + 2.0;
                                vybe_widgets::tree_view::TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, label, tx, ty, CosmicColor::rgb(20,20,20), scale);
                            }
                            ControlType::TextBox => {
                                let mut tf = vybe_widgets::TextInput::new();
                                tf.width = w; tf.height = h; tf.focused = Some(i) == selected_idx;
                                tf.paint_border(&mut pixmap, x, y, scale);
                                let label = ctrl.get_text().unwrap_or(&ctrl.name);
                                let tx = x + 6.0;
                                let ty = y + (h - 14.0) / 2.0 + 2.0;
                                vybe_widgets::tree_view::TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, label, tx, ty, CosmicColor::rgb(20,20,20), scale);
                            }
                            ControlType::Label => {
                                let label = ctrl.get_text().unwrap_or(&ctrl.name);
                                let tx = x + 4.0;
                                let ty = y + (h - 14.0) / 2.0 + 2.0;
                                vybe_widgets::tree_view::TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, label, tx, ty, CosmicColor::rgb(20,20,20), scale);
                            }
                            ControlType::CheckBox => {
                                let mut cb = vybe_widgets::Checkbox::new("");
                                cb.size = h.min(w).min(16.0);
                                cb.checked = ctrl.properties.get_bool("Checked").unwrap_or(false);
                                cb.focused = Some(i) == selected_idx;
                                cb.paint(&mut pixmap, x, y, scale);
                                let label = ctrl.get_text().unwrap_or(&ctrl.name);
                                let tx = x + cb.size + 6.0;
                                let ty = y + (h - 14.0) / 2.0 + 2.0;
                                vybe_widgets::tree_view::TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, label, tx, ty, CosmicColor::rgb(20,20,20), scale);
                            }
                            // Add more specific control renderers here as needed
                            _ => {
                                // Fallback: simple filled rect (logical coords, use scale transform)
                                let mut pb2 = PathBuilder::new();
                                pb2.push_rect(Rect::from_xywh(x, y, w, h).unwrap());
                                if let Some(pth) = pb2.finish() {
                                    let mut fill = Paint::default(); fill.set_color_rgba8(240,240,255,255);
                                    let ts = Transform::from_scale(scale, scale);
                                    pixmap.fill_path(&pth, &fill, tiny_skia::FillRule::Winding, ts, None);
                                    let mut stroke = Paint::default();
                                    if Some(i) == selected_idx { stroke.set_color_rgba8(255, 100, 40, 220); } else { stroke.set_color_rgba8(60,60,60,200); }
                                    pixmap.stroke_path(&pth, &stroke, &Stroke::default(), ts, None);
                                }
                            }
                        }
                    }

                    // Draw selection handles for selected control (logical coords, scaled)
                    if let Some(idx) = selected_idx {
                        if let Some(ctrl) = form.controls.get(idx) {
                            let b = ctrl.bounds;
                            let x = b.x as f32;
                            let y = b.y as f32;
                            let w = b.width as f32;
                            let h = b.height as f32;
                            let hs = 6.0; // handle size in logical units
                            let ts = Transform::from_scale(scale, scale);
                            let corners = [
                                (x - hs/2.0, y - hs/2.0),
                                (x + w - hs/2.0, y - hs/2.0),
                                (x - hs/2.0, y + h - hs/2.0),
                                (x + w - hs/2.0, y + h - hs/2.0),
                            ];
                            for (hx, hy) in &corners {
                                let mut pbh = PathBuilder::new();
                                pbh.push_rect(Rect::from_xywh(*hx, *hy, hs, hs).unwrap());
                                if let Some(pth) = pbh.finish() {
                                    let mut fill = Paint::default(); fill.set_color_rgba8(255,100,40,220);
                                    pixmap.fill_path(&pth, &fill, tiny_skia::FillRule::Winding, ts, None);
                                    let mut stroke = Paint::default(); stroke.set_color_rgba8(0,0,0,120);
                                    pixmap.stroke_path(&pth, &stroke, &Stroke::default(), ts, None);
                                }
                            }
                        }
                    }

                    // Blit to window (softbuffer expects native-endian 0xAARRGGBB) — convert
                    let mut buffer = surface.buffer_mut().unwrap();
                    for (i, px) in pixmap.pixels().iter().enumerate() {
                        buffer[i] = (px.alpha() as u32) << 24 | (px.red() as u32) << 16 | (px.green() as u32) << 8 | (px.blue() as u32);
                    }
                    buffer.present().unwrap();
                }
                _ => {}
            },
            _ => {}
        }
    });
}
