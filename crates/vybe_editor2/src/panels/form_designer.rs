use egui::{Ui, Rect, Vec2, Pos2, Color32, Stroke};
use uuid::Uuid;
use vybe_forms::ControlType;
use crate::state::{EditorState, DragState, LassoState};

const GRID: f32 = 8.0;

fn snap(v: f32) -> f32 { (v / GRID).round() * GRID }

// ── Resize handle indices ─────────────────────────────────────────────────────
//  0=TL  1=T   2=TR
//  3=L         4=R
//  5=BL  6=B   7=BR
fn handle_positions(rect: Rect) -> [Pos2; 8] {
    let c = rect.center();
    [
        rect.left_top(),                     Pos2::new(c.x, rect.top()),  rect.right_top(),
        Pos2::new(rect.left(), c.y),                                       Pos2::new(rect.right(), c.y),
        rect.left_bottom(),                  Pos2::new(c.x, rect.bottom()), rect.right_bottom(),
    ]
}

pub fn show(ui: &mut Ui, state: &mut EditorState) {
    let Some(form) = state.current_form_data() else {
        ui.centered_and_justified(|ui| { ui.label("No form selected."); });
        return;
    };

    let form_w = form.width as f32;
    let form_h = form.height as f32;
    let form_title = form.text.clone();
    let form_back = form.back_color.clone()
        .and_then(|s| parse_hex_color(&s))
        .unwrap_or(Color32::from_rgb(240, 240, 240));

    // Snapshot all control data (visual + non-visual)
    let controls: Vec<(Uuid, String, ControlType, f32, f32, f32, f32, String)> =
        form.controls.iter().map(|c| (
            c.id,
            c.name.clone(),
            c.control_type.clone(),
            c.bounds.x as f32,
            c.bounds.y as f32,
            c.bounds.width as f32,
            c.bounds.height as f32,
            c.properties.get_string("Text").unwrap_or_default().to_string(),
        )).collect();

    egui::ScrollArea::both().show(ui, |ui| {
        // ── Title bar ─────────────────────────────────────────────────────
        let title_rect = ui.allocate_exact_size(
            Vec2::new(form_w, 24.0),
            egui::Sense::hover(),
        ).0;
        ui.painter().rect_filled(title_rect, 0.0, Color32::from_rgb(0, 84, 166));
        ui.painter().text(
            title_rect.center(),
            egui::Align2::CENTER_CENTER,
            &form_title,
            egui::FontId::proportional(13.0),
            Color32::WHITE,
        );

        // ── Form canvas ───────────────────────────────────────────────────
        let (canvas_resp, painter) = ui.allocate_painter(
            Vec2::new(form_w, form_h),
            egui::Sense::click_and_drag(),
        );
        let origin = canvas_resp.rect.min;

        // Background + dot grid
        painter.rect_filled(canvas_resp.rect, 0.0, form_back);
        let dot_color = Color32::from_rgba_premultiplied(140, 140, 140, 90);
        let mut gx = 0.0f32;
        while gx <= form_w {
            let mut gy = 0.0f32;
            while gy <= form_h {
                painter.circle_filled(origin + Vec2::new(gx, gy), 0.8, dot_color);
                gy += GRID;
            }
            gx += GRID;
        }

        // ── Canvas interactions ───────────────────────────────────────────

        // Drag start on canvas background → begin lasso
        if canvas_resp.drag_started() && state.selected_tool.is_none() {
            if let Some(pos) = canvas_resp.interact_pointer_pos() {
                let local = pos - origin;
                state.lasso = Some(LassoState {
                    origin: local.to_pos2(),
                    current: local.to_pos2(),
                });
            }
        }

        // Drag ongoing → update lasso endpoint
        if canvas_resp.dragged() {
            if let Some(ref mut ls) = state.lasso {
                if let Some(pos) = canvas_resp.interact_pointer_pos() {
                    ls.current = (pos - origin).to_pos2();
                }
            }
        }

        // Drag released → finish lasso selection
        if canvas_resp.drag_stopped() {
            if let Some(ls) = state.lasso.take() {
                let lx = ls.origin.x.min(ls.current.x);
                let ly = ls.origin.y.min(ls.current.y);
                let lw = (ls.origin.x - ls.current.x).abs();
                let lh = (ls.origin.y - ls.current.y).abs();
                if lw > 4.0 || lh > 4.0 {
                    let mut hits = Vec::new();
                    for (id, _, ct, cx, cy, cw, ch, _) in &controls {
                        if ct.is_non_visual() { continue; }
                        // AABB intersection
                        if cx < &(lx + lw) && &(cx + cw) > &lx && cy < &(ly + lh) && &(cy + ch) > &ly {
                            hits.push(*id);
                        }
                    }
                    state.selected_controls = hits;
                }
            }
        }

        // Click on canvas background
        if canvas_resp.clicked() {
            if let Some(tool) = state.selected_tool.clone() {
                if let Some(pos) = canvas_resp.interact_pointer_pos() {
                    let lx = snap(pos.x - origin.x) as i32;
                    let ly = snap(pos.y - origin.y) as i32;
                    state.add_control(tool, lx, ly);
                    state.selected_tool = None;
                }
            } else if state.lasso.is_none() {
                state.selected_controls.clear();
            }
        }

        // ── Draw visual controls ──────────────────────────────────────────
        // Collect which IDs are selected before any mutable borrows
        let selected_ids: Vec<Uuid> = state.selected_controls.clone();

        // Snapshot drag state for use in drawing
        let drag_snap = state.drag.clone();

        for (id, name, ct, cx, cy, cw, ch, text) in &controls {
            if ct.is_non_visual() { continue; }

            let rect = Rect::from_min_size(
                origin + Vec2::new(*cx, *cy),
                Vec2::new(*cw, *ch),
            );

            let is_selected = selected_ids.contains(id);
            let fill = control_fill(ct);
            painter.rect_filled(rect, 2.0, fill);
            painter.rect_stroke(
                rect, 2.0,
                Stroke::new(
                    if is_selected { 2.0 } else { 1.0 },
                    if is_selected { Color32::from_rgb(0, 120, 212) } else { Color32::DARK_GRAY },
                ),
                egui::StrokeKind::Outside,
            );

            // Rich control visuals
            draw_control_content(&painter, rect, ct, text, name);

            // Selection handles (single-selection only)
            if is_selected && selected_ids.len() == 1 {
                for hp in handle_positions(rect) {
                    let hr = Rect::from_center_size(hp, Vec2::splat(6.0));
                    painter.rect_filled(hr, 1.0, Color32::WHITE);
                    painter.rect_stroke(hr, 1.0, Stroke::new(1.0, Color32::from_rgb(0, 120, 212)), egui::StrokeKind::Outside);
                }
            }

            // Control interaction
            let ctrl_resp = ui.interact(rect, egui::Id::new(id), egui::Sense::click_and_drag());

            if ctrl_resp.clicked() && state.lasso.is_none() {
                state.selected_controls = vec![*id];
            }

            // Drag start: snapshot all selected bounds for multi-move
            if ctrl_resp.drag_started() {
                state.push_undo();
                if !selected_ids.contains(id) {
                    state.selected_controls = vec![*id];
                }
                // Build initial-bounds snapshot for all selected controls
                if let Some(form) = state.current_form_data() {
                    let sel_now = state.selected_controls.clone();
                    let initial_bounds: Vec<(Uuid, vybe_forms::Bounds)> = form.controls.iter()
                        .filter(|c| sel_now.contains(&c.id))
                        .map(|c| (c.id, c.bounds))
                        .collect();
                    state.drag = Some(DragState {
                        ids: sel_now,
                        start_mouse: ctrl_resp.interact_pointer_pos().unwrap_or_default(),
                        initial_bounds,
                    });
                }
            }

            // Drag ongoing: move ALL selected controls by same delta
            if ctrl_resp.dragged() {
                if let Some(ref ds) = drag_snap {
                    let current_pos = ctrl_resp.interact_pointer_pos().unwrap_or_default();
                    let delta = current_pos - ds.start_mouse;
                    if let Some(form) = state.current_form_data_mut() {
                        for (cid, ib) in &ds.initial_bounds {
                            if let Some(c) = form.controls.iter_mut().find(|c| c.id == *cid) {
                                c.bounds.x = snap(ib.x as f32 + delta.x) as i32;
                                c.bounds.y = snap(ib.y as f32 + delta.y) as i32;
                            }
                        }
                    }
                }
            }

            if ctrl_resp.drag_stopped() {
                state.drag = None;
            }
        }

        // ── Resize handles interaction (single selection) ─────────────────
        if selected_ids.len() == 1 {
            let sel_id = selected_ids[0];
            // Find the selected visual control's rect
            if let Some((_, _, ct, cx, cy, cw, ch, _)) = controls.iter().find(|(id, _, ct, ..)| *id == sel_id && !ct.is_non_visual()) {
                let rect = Rect::from_min_size(
                    origin + Vec2::new(*cx, *cy),
                    Vec2::new(*cw, *ch),
                );
                let handles = handle_positions(rect);
                let handle_names = ["tl","t","tr","l","r","bl","b","br"];

                for (hi, (hp, hname)) in handles.iter().zip(handle_names.iter()).enumerate() {
                    let hr = Rect::from_center_size(*hp, Vec2::splat(10.0));
                    let hresp = ui.interact(hr, egui::Id::new((sel_id, *hname)), egui::Sense::drag());

                    if hresp.drag_started() {
                        state.push_undo();
                    }

                    if hresp.dragged() {
                        let d = hresp.drag_delta();
                        if let Some(form) = state.current_form_data_mut() {
                            if let Some(ctrl) = form.controls.iter_mut().find(|c| c.id == sel_id) {
                                let b = &mut ctrl.bounds;
                                match hi {
                                    0 => { // TL
                                        b.x = snap(b.x as f32 + d.x) as i32;
                                        b.y = snap(b.y as f32 + d.y) as i32;
                                        b.width  = (b.width  - d.x as i32).max(16);
                                        b.height = (b.height - d.y as i32).max(16);
                                    }
                                    1 => { // T
                                        b.y = snap(b.y as f32 + d.y) as i32;
                                        b.height = (b.height - d.y as i32).max(16);
                                    }
                                    2 => { // TR
                                        b.y = snap(b.y as f32 + d.y) as i32;
                                        b.width  = (b.width  + d.x as i32).max(16);
                                        b.height = (b.height - d.y as i32).max(16);
                                    }
                                    3 => { // L
                                        b.x = snap(b.x as f32 + d.x) as i32;
                                        b.width = (b.width - d.x as i32).max(16);
                                    }
                                    4 => { // R
                                        b.width = snap(b.width as f32 + d.x) as i32;
                                        b.width = b.width.max(16);
                                    }
                                    5 => { // BL
                                        b.x = snap(b.x as f32 + d.x) as i32;
                                        b.width  = (b.width  - d.x as i32).max(16);
                                        b.height = snap(b.height as f32 + d.y) as i32;
                                        b.height = b.height.max(16);
                                    }
                                    6 => { // B
                                        b.height = snap(b.height as f32 + d.y) as i32;
                                        b.height = b.height.max(16);
                                    }
                                    7 => { // BR
                                        b.width  = snap(b.width  as f32 + d.x) as i32;
                                        b.height = snap(b.height as f32 + d.y) as i32;
                                        b.width  = b.width.max(16);
                                        b.height = b.height.max(16);
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                }
            }
        }

        // ── Lasso overlay ─────────────────────────────────────────────────
        if let Some(ref ls) = state.lasso {
            let lx = ls.origin.x.min(ls.current.x);
            let ly = ls.origin.y.min(ls.current.y);
            let lw = (ls.origin.x - ls.current.x).abs();
            let lh = (ls.origin.y - ls.current.y).abs();
            let lasso_rect = Rect::from_min_size(origin + Vec2::new(lx, ly), Vec2::new(lw, lh));
            painter.rect_filled(lasso_rect, 0.0, Color32::from_rgba_premultiplied(0, 120, 212, 22));
            painter.rect_stroke(lasso_rect, 0.0, Stroke::new(1.0, Color32::from_rgb(0, 120, 212)), egui::StrokeKind::Outside);
        }

        // ── Keyboard shortcuts ────────────────────────────────────────────
        if canvas_resp.hovered() || canvas_resp.has_focus() {
            let ctx = ui.ctx();
            if ctx.input(|i| i.key_pressed(egui::Key::Delete) || i.key_pressed(egui::Key::Backspace)) {
                state.delete_selected();
            }
            if ctx.input(|i| i.modifiers.ctrl && i.key_pressed(egui::Key::Z)) {
                state.undo();
            }
            if ctx.input(|i| i.modifiers.ctrl && (i.key_pressed(egui::Key::Y) || (i.modifiers.shift && i.key_pressed(egui::Key::Z)))) {
                state.redo();
            }
            if ctx.input(|i| i.modifiers.ctrl && i.key_pressed(egui::Key::A)) {
                state.selected_controls = controls.iter()
                    .filter(|(_, _, ct, ..)| !ct.is_non_visual())
                    .map(|(id, ..)| *id)
                    .collect();
            }
            if ctx.input(|i| i.modifiers.ctrl && i.key_pressed(egui::Key::C)) {
                state.copy_selected();
            }
            if ctx.input(|i| i.modifiers.ctrl && i.key_pressed(egui::Key::X)) {
                state.cut_selected();
            }
            if ctx.input(|i| i.modifiers.ctrl && i.key_pressed(egui::Key::V)) {
                state.paste();
            }
        }

        // ── Component Tray (non-visual controls) ──────────────────────────
        let non_visual: Vec<_> = controls.iter()
            .filter(|(_, _, ct, ..)| ct.is_non_visual())
            .collect();

        if !non_visual.is_empty() {
            // Divider
            ui.add_space(4.0);
            let tray_sep = ui.allocate_exact_size(Vec2::new(form_w, 2.0), egui::Sense::hover()).0;
            ui.painter().rect_filled(tray_sep, 0.0, Color32::from_rgb(180, 180, 180));

            ui.add_space(2.0);
            ui.label(egui::RichText::new("Component Tray").small().weak());
            ui.add_space(4.0);

            ui.horizontal_wrapped(|ui| {
                for (id, name, ct, ..) in &non_visual {
                    let is_sel = state.selected_controls.contains(id);
                    let icon = tray_icon(ct);
                    let (bg, border_color) = if is_sel {
                        (Color32::from_rgb(204, 228, 247), Color32::from_rgb(0, 120, 212))
                    } else {
                        (Color32::from_rgb(232, 232, 232), Color32::from_rgb(180, 180, 180))
                    };

                    // Allocate a fixed-size chip rect so chips wrap properly
                    let chip_size = Vec2::new(68.0, 52.0);
                    let (chip_rect, chip_resp) = ui.allocate_exact_size(chip_size, egui::Sense::click());

                    if ui.is_rect_visible(chip_rect) {
                        let p = ui.painter();
                        p.rect_filled(chip_rect, 4.0, bg);
                        p.rect_stroke(chip_rect, 4.0,
                            Stroke::new(if is_sel { 2.0 } else { 1.0 }, border_color),
                            egui::StrokeKind::Outside,
                        );
                        // Icon
                        p.text(
                            chip_rect.center_top() + Vec2::new(0.0, 14.0),
                            egui::Align2::CENTER_CENTER,
                            icon,
                            egui::FontId::proportional(18.0),
                            Color32::from_rgb(30, 30, 30),
                        );
                        // Name label — truncate if needed
                        let disp_name = if name.len() > 10 { &name[..10] } else { name.as_str() };
                        p.text(
                            chip_rect.center_bottom() - Vec2::new(0.0, 10.0),
                            egui::Align2::CENTER_CENTER,
                            disp_name,
                            egui::FontId::proportional(9.0),
                            Color32::from_rgb(60, 60, 60),
                        );
                    }

                    if chip_resp.clicked() {
                        state.selected_controls = vec![*id];
                    }
                }
            });
        }
    });
}

// ── Control fill colours ──────────────────────────────────────────────────────
fn control_fill(ct: &ControlType) -> Color32 {
    match ct {
        ControlType::Button            => Color32::from_rgb(225, 225, 225),
        ControlType::Label | ControlType::LinkLabel => Color32::TRANSPARENT,
        ControlType::TextBox | ControlType::RichTextBox | ControlType::MaskedTextBox => Color32::WHITE,
        ControlType::CheckBox | ControlType::RadioButton => Color32::TRANSPARENT,
        ControlType::ComboBox | ControlType::ListBox     => Color32::WHITE,
        ControlType::Panel | ControlType::Frame          => Color32::from_rgb(235, 235, 235),
        ControlType::DataGridView | ControlType::ListView => Color32::WHITE,
        ControlType::ProgressBar      => Color32::from_rgb(220, 220, 220),
        ControlType::MenuStrip | ControlType::StatusStrip | ControlType::ToolStrip => Color32::from_rgb(240, 240, 240),
        _                             => Color32::from_rgb(230, 230, 230),
    }
}

// ── Rich control rendering ────────────────────────────────────────────────────
fn draw_control_content(painter: &egui::Painter, rect: Rect, ct: &ControlType, text: &str, name: &str) {
    let label = if text.is_empty() { name } else { text };
    let font  = || egui::FontId::proportional(11.0);
    let small = || egui::FontId::proportional(10.0);
    let black = Color32::from_rgb(30, 30, 30);
    let gray  = Color32::from_rgb(120, 120, 120);
    let white = Color32::WHITE;
    let blue  = Color32::from_rgb(0, 120, 212);

    match ct {
        ControlType::Button => {
            painter.text(rect.center(), egui::Align2::CENTER_CENTER, label, font(), black);
        }
        ControlType::Label => {
            painter.text(rect.left_center() + Vec2::new(2.0, 0.0), egui::Align2::LEFT_CENTER, label, font(), black);
        }
        ControlType::LinkLabel => {
            painter.text(rect.left_center() + Vec2::new(2.0, 0.0), egui::Align2::LEFT_CENTER, label, font(), blue);
        }
        ControlType::TextBox | ControlType::MaskedTextBox => {
            painter.text(rect.left_center() + Vec2::new(4.0, 0.0), egui::Align2::LEFT_CENTER, label, font(), black);
        }
        ControlType::RichTextBox => {
            painter.text(rect.left_top() + Vec2::new(4.0, 4.0), egui::Align2::LEFT_TOP, label, font(), black);
        }
        ControlType::CheckBox => {
            let box_size = 13.0;
            let bx = rect.min.x + 3.0;
            let by = rect.center().y - box_size / 2.0;
            let box_rect = Rect::from_min_size(Pos2::new(bx, by), Vec2::splat(box_size));
            painter.rect_filled(box_rect, 1.0, white);
            painter.rect_stroke(box_rect, 1.0, Stroke::new(1.0, gray), egui::StrokeKind::Outside);
            painter.text(
                Pos2::new(bx + box_size + 4.0, rect.center().y),
                egui::Align2::LEFT_CENTER, label, font(), black,
            );
        }
        ControlType::RadioButton => {
            let r = 6.5;
            let cx = rect.min.x + 3.0 + r;
            let cy = rect.center().y;
            painter.circle_filled(Pos2::new(cx, cy), r, white);
            painter.circle_stroke(Pos2::new(cx, cy), r, Stroke::new(1.0, gray));
            painter.text(
                Pos2::new(cx + r + 4.0, cy),
                egui::Align2::LEFT_CENTER, label, font(), black,
            );
        }
        ControlType::ComboBox => {
            let btn_w = 17.0;
            let btn_rect = Rect::from_min_size(
                Pos2::new(rect.max.x - btn_w, rect.min.y),
                Vec2::new(btn_w, rect.height()),
            );
            painter.rect_filled(btn_rect, 0.0, Color32::from_rgb(220, 220, 220));
            painter.rect_stroke(btn_rect, 0.0, Stroke::new(1.0, gray), egui::StrokeKind::Outside);
            painter.text(btn_rect.center(), egui::Align2::CENTER_CENTER, "▼", small(), black);
            let txt_rect = Rect::from_min_max(rect.min, Pos2::new(rect.max.x - btn_w, rect.max.y));
            painter.text(txt_rect.left_center() + Vec2::new(4.0, 0.0), egui::Align2::LEFT_CENTER, label, font(), black);
        }
        ControlType::ListBox => {
            let items = ["(item 1)", "(item 2)", "(item 3)"];
            let row_h = (rect.height() / items.len() as f32).min(20.0);
            for (i, item) in items.iter().enumerate() {
                let row = Rect::from_min_size(
                    rect.min + Vec2::new(0.0, i as f32 * row_h),
                    Vec2::new(rect.width(), row_h),
                );
                if i == 0 {
                    painter.rect_filled(row, 0.0, blue);
                    painter.text(row.left_center() + Vec2::new(4.0, 0.0), egui::Align2::LEFT_CENTER, item, small(), white);
                } else {
                    painter.text(row.left_center() + Vec2::new(4.0, 0.0), egui::Align2::LEFT_CENTER, item, small(), black);
                }
            }
        }
        ControlType::DataGridView => {
            let header_h = 20.0;
            let header_rect = Rect::from_min_size(rect.min, Vec2::new(rect.width(), header_h));
            painter.rect_filled(header_rect, 0.0, Color32::from_rgb(240, 240, 240));
            painter.rect_stroke(header_rect, 0.0, Stroke::new(1.0, Color32::from_rgb(200, 200, 200)), egui::StrokeKind::Outside);
            let col_w = rect.width() / 3.0;
            for (i, col) in ["Column1", "Column2", "Column3"].iter().enumerate() {
                let cx = rect.min.x + i as f32 * col_w + 4.0;
                painter.text(Pos2::new(cx, header_rect.center().y), egui::Align2::LEFT_CENTER, col, small(), black);
                if i < 2 {
                    painter.line_segment(
                        [Pos2::new(rect.min.x + (i + 1) as f32 * col_w, rect.min.y),
                         Pos2::new(rect.min.x + (i + 1) as f32 * col_w, rect.min.y + header_h)],
                        Stroke::new(1.0, Color32::from_rgb(200, 200, 200)),
                    );
                }
            }
            // Empty rows hint
            painter.text(
                Pos2::new(rect.center().x, rect.min.y + header_h + (rect.height() - header_h) / 2.0),
                egui::Align2::CENTER_CENTER, "(data)", small(), gray,
            );
        }
        ControlType::TreeView => {
            let nodes = [("▶ Node 1", 0.0), ("▼ Node 2", 0.0), ("  ▶ Child 1", 16.0), ("    Child 2",  32.0)];
            let row_h = 16.0;
            for (i, (node, indent)) in nodes.iter().enumerate() {
                let p = Pos2::new(rect.min.x + 4.0 + indent, rect.min.y + 4.0 + i as f32 * row_h);
                if rect.contains(p) {
                    painter.text(p, egui::Align2::LEFT_TOP, node, small(), black);
                }
            }
        }
        ControlType::ListView => {
            let header_h = 18.0;
            let header_rect = Rect::from_min_size(rect.min, Vec2::new(rect.width(), header_h));
            painter.rect_filled(header_rect, 0.0, Color32::from_rgb(240, 240, 240));
            let col_w = rect.width() / 3.0;
            for (i, col) in ["Name", "Type", "Size"].iter().enumerate() {
                let cx = rect.min.x + i as f32 * col_w + 4.0;
                painter.text(Pos2::new(cx, header_rect.center().y), egui::Align2::LEFT_CENTER, col, small(), black);
            }
            painter.text(
                Pos2::new(rect.center().x, rect.min.y + header_h + 8.0),
                egui::Align2::CENTER_CENTER, "(empty)", small(), gray,
            );
        }
        ControlType::ProgressBar => {
            let fill_w = rect.width() * 0.3;
            let fill_rect = Rect::from_min_size(rect.min, Vec2::new(fill_w, rect.height()));
            painter.rect_filled(fill_rect, 0.0, Color32::from_rgb(6, 176, 37));
        }
        ControlType::TrackBar => {
            let cy = rect.center().y;
            let track_rect = Rect::from_center_size(
                Pos2::new(rect.center().x, cy),
                Vec2::new(rect.width() - 10.0, 4.0),
            );
            painter.rect_filled(track_rect, 2.0, Color32::from_rgb(200, 200, 200));
            painter.rect_stroke(track_rect, 2.0, Stroke::new(1.0, gray), egui::StrokeKind::Outside);
            // Thumb at 30%
            let thumb_x = track_rect.min.x + track_rect.width() * 0.3;
            let thumb = Rect::from_center_size(Pos2::new(thumb_x, cy), Vec2::new(8.0, 16.0));
            painter.rect_filled(thumb, 2.0, Color32::from_rgb(80, 80, 80));
        }
        ControlType::MenuStrip | ControlType::ToolStrip | ControlType::StatusStrip => {
            let items = match ct {
                ControlType::MenuStrip => vec!["File", "Edit", "View", "Help"],
                ControlType::ToolStrip => vec!["✂", "📋", "📌", "|", "▶"],
                _                       => vec!["Ready"],
            };
            let mut x = rect.min.x + 4.0;
            for item in items {
                painter.text(Pos2::new(x, rect.center().y), egui::Align2::LEFT_CENTER, item, font(), black);
                x += item.len() as f32 * 7.0 + 12.0;
            }
        }
        ControlType::Panel | ControlType::Frame => {
            if matches!(ct, ControlType::Frame) {
                painter.text(
                    rect.left_top() + Vec2::new(6.0, -6.0),
                    egui::Align2::LEFT_CENTER, label, small(), black,
                );
            }
        }
        ControlType::TabControl => {
            let tab_h = 22.0;
            let tabs = ["Tab 1", "Tab 2"];
            let tab_w = rect.width() / tabs.len() as f32;
            for (i, tab) in tabs.iter().enumerate() {
                let tab_rect = Rect::from_min_size(
                    rect.min + Vec2::new(i as f32 * tab_w, 0.0),
                    Vec2::new(tab_w, tab_h),
                );
                let tab_fill = if i == 0 { white } else { Color32::from_rgb(220, 220, 220) };
                painter.rect_filled(tab_rect, 0.0, tab_fill);
                painter.rect_stroke(tab_rect, 0.0, Stroke::new(1.0, gray), egui::StrokeKind::Outside);
                painter.text(tab_rect.center(), egui::Align2::CENTER_CENTER, tab, small(), black);
            }
        }
        _ => {
            // Generic fallback: render the label
            painter.text(rect.center(), egui::Align2::CENTER_CENTER, label, font(), black);
        }
    }
}

// ── Helper: tray icon for a non-visual control type ──────────────────────────
fn tray_icon(ct: &ControlType) -> &'static str {
    match ct {
        ControlType::BindingSourceComponent  => "🔗",
        ControlType::DataSetComponent        => "🗄",
        ControlType::DataTableComponent      => "📋",
        ControlType::DataAdapterComponent    => "🔌",
        ControlType::BindingNavigator        => "🧭",
        ControlType::Timer                   => "⏱",
        ControlType::ImageList               => "🖼",
        ControlType::ErrorProvider           => "⚠",
        ControlType::BackgroundWorker        => "⚙",
        ControlType::OpenFileDialog          => "📂",
        ControlType::SaveFileDialog          => "💾",
        ControlType::FolderBrowserDialog     => "📁",
        ControlType::FontDialog              => "🔤",
        ControlType::ColorDialog             => "🎨",
        ControlType::NotifyIcon              => "🔔",
        ControlType::PrintDialog |
        ControlType::PrintDocument           => "🖨",
        _                                    => "📦",
    }
}

// ── Helper: parse "#rrggbb" → Color32 ────────────────────────────────────────
fn parse_hex_color(s: &str) -> Option<Color32> {
    let s = s.trim_start_matches('#');
    if s.len() == 6 {
        let r = u8::from_str_radix(&s[0..2], 16).ok()?;
        let g = u8::from_str_radix(&s[2..4], 16).ok()?;
        let b = u8::from_str_radix(&s[4..6], 16).ok()?;
        Some(Color32::from_rgb(r, g, b))
    } else {
        None
    }
}
