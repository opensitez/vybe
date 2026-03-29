use eframe::egui;
use std::path::Path;
use vybe_project::resources::ResourceType;
use vybe_project::ResourceItem;
use crate::state::EditorState;

#[derive(Clone)]
struct LocalState {
    active_category: ResourceType,
    new_name: String,
    new_value: String,
    new_comment: String,
}

impl Default for LocalState {
    fn default() -> Self {
        Self {
            active_category: ResourceType::String,
            new_name: String::new(),
            new_value: String::new(),
            new_comment: String::new(),
        }
    }
}

fn name_from_path(path: &str) -> String {
    Path::new(path)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Resource1")
        .replace(|c: char| !c.is_alphanumeric() && c != '_', "_")
        .to_string()
}

fn pick_files_for_type(rt: &ResourceType) -> Vec<(String, String)> {
    let mut dialog = rfd::FileDialog::new();
    let exts = match rt {
        ResourceType::Image => vec!["png", "jpg", "jpeg", "gif", "bmp", "tiff", "webp"],
        ResourceType::Icon => vec!["ico"],
        ResourceType::Audio => vec!["wav", "mp3", "ogg", "flac", "aiff"],
        _ => vec![],
    };
    if !exts.is_empty() {
        dialog = dialog.add_filter(match rt {
            ResourceType::Image => "Images",
            ResourceType::Icon => "Icons",
            ResourceType::Audio => "Audio",
            _ => "",
        }, &exts);
    }
    dialog = dialog.add_filter("All Files", &["*"]);
    
    if let Some(paths) = dialog.pick_files() {
        paths.iter().filter_map(|p| {
            let path_str = p.to_string_lossy().to_string();
            let name = name_from_path(&path_str);
            Some((name, path_str))
        }).collect()
    } else {
        vec![]
    }
}

fn pick_single_file_for_type(rt: &ResourceType) -> Option<(String, String)> {
    let mut dialog = rfd::FileDialog::new();
    let exts = match rt {
        ResourceType::Image => vec!["png", "jpg", "jpeg", "gif", "bmp", "tiff", "webp"],
        ResourceType::Icon => vec!["ico"],
        ResourceType::Audio => vec!["wav", "mp3", "ogg", "flac", "aiff"],
        _ => vec![],
    };
    if !exts.is_empty() {
        dialog = dialog.add_filter("Supported", &exts);
    }
    dialog.pick_file().map(|p| {
        let path_str = p.to_string_lossy().to_string();
        let name = name_from_path(&path_str);
        (name, path_str)
    })
}

fn is_file_category(rt: &ResourceType) -> bool {
    matches!(rt, ResourceType::Image | ResourceType::Icon | ResourceType::Audio | ResourceType::File)
}

pub fn show(ui: &mut egui::Ui, state: &mut EditorState) {
    let local_id = egui::Id::new("resource_editor_local");
    let mut local = ui.ctx().data_mut(|d| d.get_temp::<LocalState>(local_id).unwrap_or_default());

    if state.project.is_none() {
        ui.centered_and_justified(|ui| { ui.label("No project loaded."); });
        return;
    }
    
    // Auto-create resource file if missing
    let mut did_add = false;
    if let Some(proj) = state.project.as_mut() {
        if proj.resource_files.is_empty() {
            proj.resource_files.push(vybe_project::ResourceManager::new());
            did_add = true;
        }
    }
    if did_add {
        // Fall back or safe update? Need mutability separation.
    }
    
    let is_file_cat = is_file_category(&local.active_category);

    ui.vertical(|ui| {
        // Top Toolbar
        ui.horizontal(|ui| {
            ui.heading("Resource Editor");
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if is_file_cat {
                    if ui.button(format!("Add {}...", local.active_category.to_string())).clicked() {
                        let picked = pick_files_for_type(&local.active_category);
                        if let Some(proj) = state.project.as_mut() {
                            if let Some(rm) = proj.resource_files.first_mut() {
                                for (name, path) in picked {
                                    rm.resources.push(ResourceItem::new_file(name, &path, local.active_category.clone()));
                                }
                            }
                        }
                    }
                }
            });
        });
        ui.add_space(8.0);
        
        // Tabs
        ui.horizontal(|ui| {
            let tabs = [
                (ResourceType::String, "Strings"),
                (ResourceType::Image, "Images"),
                (ResourceType::Icon, "Icons"),
                (ResourceType::Audio, "Audio"),
                (ResourceType::File, "Files"),
                (ResourceType::Other, "Other"),
            ];
            for (rt, label) in tabs {
                if ui.selectable_value(&mut local.active_category, rt, label).changed() {
                    local.new_name.clear();
                    local.new_value.clear();
                    local.new_comment.clear();
                }
            }
        });
        ui.separator();

        let mut to_remove: Option<usize> = None;

        // Content Table
        egui::ScrollArea::both().auto_shrink([false, false]).max_height(ui.available_height() - 40.0).show(ui, |ui| {
            if let Some(proj) = state.project.as_mut() {
                if let Some(rm) = proj.resource_files.first_mut() {
                    egui::Grid::new("resource_grid").num_columns(4).spacing([16.0, 8.0]).striped(true).show(ui, |ui| {
                        ui.label(egui::RichText::new("Name").strong());
                        ui.label(egui::RichText::new(if is_file_cat { "File Path" } else { "Value" }).strong());
                        ui.label(egui::RichText::new("Comment").strong());
                        ui.label(egui::RichText::new("Actions").strong());
                        ui.end_row();

                        let mut index = 0;
                        for res in rm.resources.iter_mut() {
                            if res.resource_type == local.active_category {
                                ui.text_edit_singleline(&mut res.name);
                                
                                if is_file_category(&res.resource_type) {
                                    ui.horizontal(|ui| {
                                        let file_name = Path::new(&res.value).file_name().and_then(|s| s.to_str()).unwrap_or(&res.value).to_string();
                                        ui.label(file_name).on_hover_text(&res.value);
                                        if ui.small_button("Browse").clicked() {
                                            if let Some((_, p)) = pick_single_file_for_type(&res.resource_type) {
                                                res.value = p.clone();
                                                res.file_name = Some(p);
                                            }
                                        }
                                    });
                                } else {
                                    ui.text_edit_singleline(&mut res.value);
                                }
                                
                                let mut comment = res.comment.clone().unwrap_or_default();
                                if ui.text_edit_singleline(&mut comment).changed() {
                                    res.comment = Some(comment);
                                }
                                
                                if ui.button("❌").clicked() {
                                    to_remove = Some(index);
                                }
                                ui.end_row();
                            }
                            index += 1;
                        }
                    });
                    
                    if let Some(idx) = to_remove {
                        rm.resources.remove(idx);
                    }
                }
            }
        });

        // Bottom Add Row for Strings/Other
        if !is_file_cat {
            ui.separator();
            ui.horizontal(|ui| {
                ui.add(egui::TextEdit::singleline(&mut local.new_name).hint_text("Name").desired_width(120.0));
                ui.add(egui::TextEdit::singleline(&mut local.new_value).hint_text("Value").desired_width(200.0));
                ui.add(egui::TextEdit::singleline(&mut local.new_comment).hint_text("Comment (optional)").desired_width(120.0));
                
                if ui.button(format!("Add {}", local.active_category.to_string())).clicked() {
                    if !local.new_name.is_empty() {
                        if let Some(proj) = state.project.as_mut() {
                            if let Some(rm) = proj.resource_files.first_mut() {
                                let mut item = ResourceItem::new_string(local.new_name.clone(), local.new_value.clone());
                                item.resource_type = local.active_category.clone();
                                item.comment = if local.new_comment.is_empty() { None } else { Some(local.new_comment.clone()) };
                                rm.resources.push(item);
                                local.new_name.clear();
                                local.new_value.clear();
                                local.new_comment.clear();
                            }
                        }
                    }
                }
            });
        }
    });

    ui.ctx().data_mut(|d| d.insert_temp(local_id, local));
}
