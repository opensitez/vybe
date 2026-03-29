use egui::{Ui, Rect, Vec2, Pos2, Color32, Stroke, Response};
use crate::state::EditorState;

const GRID: f32 = 8.0;

fn snap(v: f32) -> f32 { (v / GRID).round() * GRID }

pub fn show(ui: &mut Ui, state: &mut EditorState) {
    let Some(form) = state.current_form_data() else {
        ui.centered_and_justified(|ui| { ui.label("No form selected."); });
        return;
    };

    let form_w = form.width as f32;
    let form_h = form.height as f32;
    let form_title = form.text.clone();
    let controls: Vec<_> = form.controls.iter().map(|c| (
        c.id,
        c.name.clone(),
        format!("{:?}", c.control_type),
        c.bounds.x as f32,
        c.bounds.y as f32,
        c.bounds.width as f32,
        c.bounds.height as f32,
        c.properties.get_string("Text").unwrap_or_default().to_string(),
    )).collect();

    egui::ScrollArea::both().show(ui, |ui| {
        // Title bar
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

        // Form canvas
        let (canvas_resp, painter) = ui.allocate_painter(
            Vec2::new(form_w, form_h),
            egui::Sense::click_and_drag(),
        );
        let origin = canvas_resp.rect.min;

        // Background + grid
        painter.rect_filled(canvas_resp.rect, 0.0, Color32::from_rgb(240, 240, 240));
        let grid_color = Color32::from_rgba_premultiplied(180, 180, 180, 80);
        let mut gx = 0.0f32;
        while gx <= form_w {
            painter.line_segment([origin + Vec2::new(gx, 0.0), origin + Vec2::new(gx, form_h)], Stroke::new(0.5, grid_color));
            gx += GRID;
        }
        let mut gy = 0.0f32;
        while gy <= form_h {
            painter.line_segment([origin + Vec2::new(0.0, gy), origin + Vec2::new(form_w, gy)], Stroke::new(0.5, grid_color));
            gy += GRID;
        }

        // Click on canvas to place control
        if canvas_resp.clicked() {
            if let Some(tool) = state.selected_tool.clone() {
                if let Some(pos) = canvas_resp.interact_pointer_pos() {
                    let lx = snap(pos.x - origin.x) as i32;
                    let ly = snap(pos.y - origin.y) as i32;
                    state.add_control(tool, lx, ly);
                    state.selected_tool = None;
                }
            } else if canvas_resp.clicked() {
                // Deselect if clicking empty canvas
                state.selected_controls.clear();
            }
        }

        // Draw controls
        for (id, name, type_name, cx, cy, cw, ch, text) in &controls {
            let rect = Rect::from_min_size(
                origin + Vec2::new(*cx, *cy),
                Vec2::new(*cw, *ch),
            );

            let is_selected = state.selected_controls.contains(id);
            let fill = control_fill(type_name);
            painter.rect_filled(rect, 2.0, fill);
            painter.rect_stroke(rect, 2.0, Stroke::new(if is_selected { 2.0 } else { 1.0 }, if is_selected { Color32::BLUE } else { Color32::DARK_GRAY }), egui::StrokeKind::Outside);

            // Label
            let label = if text.is_empty() { name.as_str() } else { text.as_str() };
            painter.text(
                rect.center(),
                egui::Align2::CENTER_CENTER,
                label,
                egui::FontId::proportional(11.0),
                Color32::BLACK,
            );

            // Selection handles
            if is_selected {
                for hp in handle_positions(rect) {
                    painter.rect_filled(
                        Rect::from_center_size(hp, Vec2::splat(6.0)),
                        0.0,
                        Color32::WHITE,
                    );
                    painter.rect_stroke(
                        Rect::from_center_size(hp, Vec2::splat(6.0)),
                        0.0,
                        Stroke::new(1.0, Color32::BLUE),
                        egui::StrokeKind::Outside,
                    );
                }
            }

            // Interaction
            let ctrl_resp = ui.interact(rect, egui::Id::new(id), egui::Sense::click_and_drag());

            if ctrl_resp.clicked() {
                state.selected_controls = vec![*id];
            }

            // Drag to move
            if ctrl_resp.dragged() {
                let delta = ctrl_resp.drag_delta();
                if let Some(form) = state.current_form_data_mut() {
                    if let Some(ctrl) = form.controls.iter_mut().find(|c| c.id == *id) {
                        ctrl.bounds.x = snap(ctrl.bounds.x as f32 + delta.x) as i32;
                        ctrl.bounds.y = snap(ctrl.bounds.y as f32 + delta.y) as i32;
                    }
                }
            }
            if ctrl_resp.drag_started() {
                state.push_undo();
                state.selected_controls = vec![*id];
            }
        }

        // Keyboard shortcuts
        if canvas_resp.hovered() {
            let ctx = ui.ctx();
            if ctx.input(|i| i.key_pressed(egui::Key::Delete)) {
                state.delete_selected();
            }
            if ctx.input(|i| i.modifiers.ctrl && i.key_pressed(egui::Key::Z)) {
                state.undo();
            }
            if ctx.input(|i| i.modifiers.ctrl && i.key_pressed(egui::Key::Y)) {
                state.redo();
            }
        }
    });
}

fn control_fill(type_name: &str) -> Color32 {
    match type_name {
        "Button" => Color32::from_rgb(225, 225, 225),
        "Label" => Color32::from_rgba_premultiplied(0, 0, 0, 0),
        "TextBox" | "RichTextBox" | "MaskedTextBox" => Color32::WHITE,
        "CheckBox" | "RadioButton" => Color32::from_rgba_premultiplied(0, 0, 0, 0),
        "ComboBox" | "ListBox" => Color32::WHITE,
        "Panel" | "Frame" => Color32::from_rgb(235, 235, 235),
        "DataGridView" => Color32::WHITE,
        _ => Color32::from_rgb(230, 230, 230),
    }
}

fn handle_positions(rect: Rect) -> [Pos2; 8] {
    let c = rect.center();
    [
        rect.left_top(), Pos2::new(c.x, rect.top()), rect.right_top(),
        Pos2::new(rect.left(), c.y),                  Pos2::new(rect.right(), c.y),
        rect.left_bottom(), Pos2::new(c.x, rect.bottom()), rect.right_bottom(),
    ]
}
