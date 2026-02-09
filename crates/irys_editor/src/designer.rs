use egui;
use irys_forms::{Control, ControlType, EventBinding, EventType};

pub fn show_designer(
    ui: &mut egui::Ui,
    project: &mut Option<irys_project::Project>,
    current_form_name: &str,
    selected_control: &mut Option<uuid::Uuid>,
    selected_tool: &mut Option<ControlType>,
    dragging_control: &mut Option<(uuid::Uuid, egui::Pos2)>,
    resizing_control: &mut Option<(uuid::Uuid, ResizeHandle)>,
    show_code_editor: &mut bool,
) {
    // Get form info for display (immutable)
    let (caption, width, height) = match project.as_ref()
        .and_then(|p| p.forms.iter().find(|fm| fm.form.name == current_form_name))
        .map(|fm| (fm.form.caption.clone(), fm.form.width, fm.form.height)) {
        Some(info) => info,
        None => {
            ui.label("No form selected");
            return;
        }
    };

    ui.horizontal(|ui| {
        ui.heading("Form Designer");
        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
            if ui.button("View Code ➡").clicked() {
                *show_code_editor = true;
            }
        });
    });

    ui.label(format!("Form: {} ({}x{})", caption, width, height));

    // Designer canvas
    let (response, painter) = ui.allocate_painter(
        egui::vec2(width as f32, height as f32),
        egui::Sense::click_and_drag(),
    );

    let rect = response.rect;

    // Draw form background
    painter.rect_filled(rect, 0.0, egui::Color32::from_gray(240));

    // Draw grid
    let grid_size = 20.0;
    for x in (0..width).step_by(grid_size as usize) {
        painter.line_segment(
            [
                egui::pos2(rect.left() + x as f32, rect.top()),
                egui::pos2(rect.left() + x as f32, rect.bottom()),
            ],
            egui::Stroke::new(0.5, egui::Color32::from_gray(220)),
        );
    }
    for y in (0..height).step_by(grid_size as usize) {
        painter.line_segment(
            [
                egui::pos2(rect.left(), rect.top() + y as f32),
                egui::pos2(rect.right(), rect.top() + y as f32),
            ],
            egui::Stroke::new(0.5, egui::Color32::from_gray(220)),
        );
    }

    // Handle dragging
    if let Some((drag_id, drag_offset)) = *dragging_control {
        if response.dragged() {
            if let Some(pos) = response.interact_pointer_pos() {
                let new_x = ((pos.x - rect.left() - drag_offset.x) / grid_size).round() * grid_size;
                let new_y = ((pos.y - rect.top() - drag_offset.y) / grid_size).round() * grid_size;

                // Update control position
                if let Some(proj) = project.as_mut() {
                    if let Some(form_module) = proj.get_form_mut(current_form_name) {
                        if let Some(control) = form_module.form.get_control_mut(drag_id) {
                            let old_x = control.bounds.x;
                            let old_y = control.bounds.y;
                            control.bounds.x = new_x as i32;
                            control.bounds.y = new_y as i32;
                            
                            // Calculate delta for moving children
                            let delta_x = new_x as i32 - old_x;
                            let delta_y = new_y as i32 - old_y;
                            
                            // If this is a frame, move all child controls
                            if control.control_type == ControlType::Frame && (delta_x != 0 || delta_y != 0) {
                                let frame_id = control.id;
                                // Collect child IDs first to avoid borrow issues
                                let child_ids: Vec<_> = form_module.form.controls.iter()
                                    .filter(|c| c.parent_id == Some(frame_id))
                                    .map(|c| c.id)
                                    .collect();
                                
                                // Move each child
                                for child_id in child_ids {
                                    if let Some(child) = form_module.form.get_control_mut(child_id) {
                                        child.bounds.x += delta_x;
                                        child.bounds.y += delta_y;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if response.drag_stopped() {
            *dragging_control = None;
        }
    }

    // Handle resizing
    if let Some((resize_id, handle)) = *resizing_control {
        if response.dragged() {
            if let Some(pos) = response.interact_pointer_pos() {
                let mouse_x = (pos.x - rect.left()) as i32;
                let mouse_y = (pos.y - rect.top()) as i32;

                // Update control size
                if let Some(proj) = project.as_mut() {
                    if let Some(form_module) = proj.get_form_mut(current_form_name) {
                        if let Some(control) = form_module.form.get_control_mut(resize_id) {
                            match handle {
                                ResizeHandle::BottomRight => {
                                    control.bounds.width = (mouse_x - control.bounds.x).max(20);
                                    control.bounds.height = (mouse_y - control.bounds.y).max(20);
                                }
                                ResizeHandle::Right => {
                                    control.bounds.width = (mouse_x - control.bounds.x).max(20);
                                }
                                ResizeHandle::Bottom => {
                                    control.bounds.height = (mouse_y - control.bounds.y).max(20);
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
        }

        if response.drag_stopped() {
            *resizing_control = None;
        }
    }

    // Get controls list for interaction (need to clone to avoid borrow issues)
    let controls = project.as_ref()
        .and_then(|p| p.forms.iter().find(|fm| fm.form.name == current_form_name))
        .map(|fm| fm.form.controls.clone())
        .unwrap_or_default();

    // Handle double-click to open event code
    if response.double_clicked() && dragging_control.is_none() && resizing_control.is_none() {
        if let Some(pos) = response.interact_pointer_pos() {
            let x = (pos.x - rect.left()) as i32;
            let y = (pos.y - rect.top()) as i32;

            // Find which control was double-clicked
            let clicked = controls.iter().find(|c| {
                x >= c.bounds.x && x <= c.bounds.x + c.bounds.width &&
                y >= c.bounds.y && y <= c.bounds.y + c.bounds.height
            }).map(|c| (c.id, c.name.clone()));

            if let Some((control_id, control_name)) = clicked {
                let handler_name = format!("{}_{}", control_name, "Click");
                let event_code = format!(
                    "Private Sub {}()\n    MsgBox \"Hello from {}!\"\nEnd Sub\n\n",
                    handler_name, control_name
                );

                // Add to code if not already there
                if let Some(proj) = project.as_mut() {
                    if let Some(form_module) = proj.get_form_mut(current_form_name) {
                        if !form_module.code.contains(&handler_name) {
                            form_module.code.push_str(&event_code);
                        }
                    }
                }

                *show_code_editor = true;
                *selected_control = Some(control_id);
            } else {
                // Double-clicked form background -> Generate Form_Load
                let handler_name = "Form_Load".to_string();
                let event_code = format!(
                    "Private Sub {}()\n    ' Initializing form {}\nEnd Sub\n\n",
                    handler_name, current_form_name
                );

                if let Some(proj) = project.as_mut() {
                    if let Some(form_module) = proj.get_form_mut(current_form_name) {
                        if !form_module.code.to_lowercase().contains("sub form_load") {
                            form_module.code.push_str(&event_code);
                        }
                    }
                }

                *show_code_editor = true;
                *selected_control = None;
            }
        }
    }

    // Handle click to place or select control
    if response.clicked() && dragging_control.is_none() && resizing_control.is_none() {
        if let Some(pos) = response.interact_pointer_pos() {
            let x = (pos.x - rect.left()) as i32;
            let y = (pos.y - rect.top()) as i32;

            if let Some(control_type) = *selected_tool {
                // Place new control
                if let Some(proj) = project.as_mut() {
                    if let Some(form_module) = proj.get_form_mut(current_form_name) {
                        let next_index = form_module
                            .form
                            .controls
                            .iter()
                            .filter(|c| c.control_type == control_type)
                            .count()
                            + 1;
                        let name = format!("{}{}", control_type.default_name_prefix(), next_index);
                        let mut control = Control::new(control_type, name.clone(), x, y);
                        
                        // Check if control is being placed inside a frame
                        for potential_parent in &form_module.form.controls {
                            if potential_parent.control_type == ControlType::Frame {
                                if x >= potential_parent.bounds.x && 
                                   x <= potential_parent.bounds.x + potential_parent.bounds.width &&
                                   y >= potential_parent.bounds.y && 
                                   y <= potential_parent.bounds.y + potential_parent.bounds.height {
                                    control.parent_id = Some(potential_parent.id);
                                    break;
                                }
                            }
                        }
                        
                        let binding = EventBinding::new(&name, EventType::Click);
                        form_module.form.add_event_binding(binding);
                        form_module.form.add_control(control);
                    }
                }
                *selected_tool = None;
            } else {
                // Select control
                let clicked_control = controls.iter()
                    .find(|c| {
                        x >= c.bounds.x && x <= c.bounds.x + c.bounds.width &&
                        y >= c.bounds.y && y <= c.bounds.y + c.bounds.height
                    })
                    .map(|c| c.id);

                *selected_control = clicked_control;
            }
        }
    }

    // Start dragging on drag start
    if response.drag_started() && selected_tool.is_none() {
        if let Some(pos) = response.interact_pointer_pos() {
            let x = (pos.x - rect.left()) as i32;
            let y = (pos.y - rect.top()) as i32;

            // Check if clicking on a resize handle first
            if let Some(sel_id) = *selected_control {
                if let Some(control) = controls.iter().find(|c| c.id == sel_id) {
                    let handle_size = 6.0;
                    let ctrl_rect = egui::Rect::from_min_size(
                        egui::pos2(control.bounds.x as f32, control.bounds.y as f32),
                        egui::vec2(control.bounds.width as f32, control.bounds.height as f32),
                    );

                    // Check resize handles
                    let handles = [
                        (ResizeHandle::BottomRight, egui::pos2(ctrl_rect.right(), ctrl_rect.bottom())),
                        (ResizeHandle::Right, egui::pos2(ctrl_rect.right(), ctrl_rect.center().y)),
                        (ResizeHandle::Bottom, egui::pos2(ctrl_rect.center().x, ctrl_rect.bottom())),
                    ];

                    let mut found_handle = false;
                    for (handle, pos_abs) in handles {
                        let handle_rect = egui::Rect::from_center_size(
                            egui::pos2(rect.left() + pos_abs.x, rect.top() + pos_abs.y),
                            egui::vec2(handle_size * 2.0, handle_size * 2.0)
                        );
                        if handle_rect.contains(egui::pos2(pos.x, pos.y)) {
                            *resizing_control = Some((sel_id, handle));
                            found_handle = true;
                            break;
                        }
                    }

                    if !found_handle && ctrl_rect.contains(egui::pos2(x as f32, y as f32)) {
                        // Start dragging the control
                        let offset = egui::pos2(
                            x as f32 - control.bounds.x as f32,
                            y as f32 - control.bounds.y as f32,
                        );
                        *dragging_control = Some((sel_id, offset));
                    }
                }
            }
        }
    }

    // Draw controls
    for control in &controls {
        let is_selected = *selected_control == Some(control.id);
        let ctrl_rect = egui::Rect::from_min_size(
            egui::pos2(
                rect.left() + control.bounds.x as f32,
                rect.top() + control.bounds.y as f32,
            ),
            egui::vec2(control.bounds.width as f32, control.bounds.height as f32),
        );

        // Background based on control type
        let bg_color = match control.control_type {
            ControlType::Button => egui::Color32::from_gray(192),
            ControlType::TextBox => egui::Color32::WHITE,
            _ => egui::Color32::from_gray(212),
        };

        painter.rect_filled(ctrl_rect, 2.0, bg_color);

        // Border
        painter.rect_stroke(
            ctrl_rect,
            2.0,
            egui::Stroke::new(1.0, egui::Color32::BLACK),
        );

        // Selection border
        if is_selected {
            painter.rect_stroke(
                ctrl_rect,
                2.0,
                egui::Stroke::new(2.0, egui::Color32::from_rgb(0, 100, 255)),
            );

            // Draw resize handles
            let handle_size = 6.0;
            let handles = [
                ctrl_rect.right_bottom(),
                egui::pos2(ctrl_rect.right(), ctrl_rect.center().y),
                egui::pos2(ctrl_rect.center().x, ctrl_rect.bottom()),
            ];

            for handle_pos in handles {
                painter.rect_filled(
                    egui::Rect::from_center_size(handle_pos, egui::vec2(handle_size, handle_size)),
                    0.0,
                    egui::Color32::WHITE,
                );
                painter.rect_stroke(
                    egui::Rect::from_center_size(handle_pos, egui::vec2(handle_size, handle_size)),
                    0.0,
                    egui::Stroke::new(1.0, egui::Color32::BLACK),
                );
            }
        }

        // Text
        let text = control.get_caption()
            .or_else(|| control.get_text())
            .unwrap_or(&control.name);

        painter.text(
            ctrl_rect.center(),
            egui::Align2::CENTER_CENTER,
            text,
            egui::FontId::default(),
            egui::Color32::BLACK,
        );
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ResizeHandle {
    TopLeft,
    TopRight,
    BottomLeft,
    BottomRight,
    Top,
    Bottom,
    Left,
    Right,
}
