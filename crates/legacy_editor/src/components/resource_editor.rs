
use dioxus::prelude::*;
use vybe_project::ResourceManager;
use vybe_project::ResourceItem;
use vybe_project::resources::ResourceType;
use std::path::Path;

/// Returns file dialog filters (description, extensions) for file-based resource types.
fn filters_for_type(rt: &ResourceType) -> Vec<(&'static str, &'static [&'static str])> {
    match rt {
        ResourceType::Image => vec![
            ("Images", &["png", "jpg", "jpeg", "gif", "bmp", "tiff", "webp"]),
            ("All Files", &["*"]),
        ],
        ResourceType::Icon => vec![
            ("Icons", &["ico"]),
            ("All Files", &["*"]),
        ],
        ResourceType::Audio => vec![
            ("Audio Files", &["wav", "mp3", "ogg", "flac", "aiff"]),
            ("All Files", &["*"]),
        ],
        ResourceType::File => vec![
            ("All Files", &["*"]),
        ],
        _ => vec![],
    }
}

/// Returns true if the category is file-based (uses a file picker instead of text input).
fn is_file_category(rt: &ResourceType) -> bool {
    matches!(rt, ResourceType::Image | ResourceType::Icon | ResourceType::Audio | ResourceType::File)
}

/// Derive a resource name from a file path (filename without extension, PascalCase-ish).
fn name_from_path(path: &str) -> String {
    Path::new(path)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Resource1")
        .replace(|c: char| !c.is_alphanumeric() && c != '_', "_")
        .to_string()
}

/// Open a native file picker for the given resource type. Returns Vec of (name, path) pairs.
fn pick_files_for_type(rt: &ResourceType) -> Vec<(String, String)> {
    let mut dialog = rfd::FileDialog::new();
    for (desc, exts) in filters_for_type(rt) {
        dialog = dialog.add_filter(desc, exts);
    }
    let title = match rt {
        ResourceType::Image => "Add Image Resource",
        ResourceType::Icon => "Add Icon Resource",
        ResourceType::Audio => "Add Audio Resource",
        ResourceType::File => "Add File Resource",
        _ => "Add Resource",
    };
    dialog = dialog.set_title(title);

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

/// Open a native file picker for replacing a single file resource. Returns (name, path) or None.
fn pick_single_file_for_type(rt: &ResourceType) -> Option<(String, String)> {
    let mut dialog = rfd::FileDialog::new();
    for (desc, exts) in filters_for_type(rt) {
        dialog = dialog.add_filter(desc, exts);
    }
    dialog = dialog.set_title("Change File");
    dialog.pick_file().map(|p| {
        let path_str = p.to_string_lossy().to_string();
        let name = name_from_path(&path_str);
        (name, path_str)
    })
}

#[derive(Props, PartialEq, Clone)]
pub struct ResourceEditorProps {
    pub resources: ResourceManager,
    pub on_change: EventHandler<ResourceManager>,
}

#[allow(non_snake_case)]
pub fn ResourceEditor(props: ResourceEditorProps) -> Element {
    let mut new_res_name = use_signal(|| String::new());
    let mut new_res_value = use_signal(|| String::new());
    let mut new_res_comment = use_signal(|| String::new());
    let mut active_category = use_signal(|| ResourceType::String);

    // Filter resources by current category
    let filtered: Vec<(usize, &ResourceItem)> = props.resources.resources.iter()
        .enumerate()
        .filter(|(_, r)| r.resource_type == *active_category.read())
        .collect();

    // Count per category
    let count_strings = props.resources.resources.iter().filter(|r| r.resource_type == ResourceType::String).count();
    let count_images = props.resources.resources.iter().filter(|r| r.resource_type == ResourceType::Image).count();
    let count_icons = props.resources.resources.iter().filter(|r| r.resource_type == ResourceType::Icon).count();
    let count_audio = props.resources.resources.iter().filter(|r| r.resource_type == ResourceType::Audio).count();
    let count_files = props.resources.resources.iter().filter(|r| r.resource_type == ResourceType::File).count();
    let count_other = props.resources.resources.iter().filter(|r| r.resource_type == ResourceType::Other).count();

    let cat = active_category.read().clone();
    let is_file_cat = is_file_category(&cat);

    let tab_style = "padding: 6px 14px; cursor: pointer; border: none; border-bottom: 2px solid transparent; background: transparent; font-size: 12px;";
    let tab_active_style = "padding: 6px 14px; cursor: pointer; border: none; border-bottom: 2px solid #0078d4; background: #e8f0fe; font-size: 12px; font-weight: bold;";

    rsx! {
        div {
            class: "resource-editor",
            style: "display: flex; flex-direction: column; height: 100%; border: 1px solid #ccc; background: white;",
            
            // Header
            div {
                style: "padding: 10px; background: #f0f0f0; border-bottom: 1px solid #ccc; display: flex; justify-content: space-between; align-items: center;",
                span { style: "font-weight: bold;", "{props.resources.name}.resx" }
                // Top-right Add button for file-based categories
                if is_file_cat {
                    button {
                        style: "padding: 4px 14px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;",
                        onclick: {
                            let resources = props.resources.clone();
                            let on_change = props.on_change;
                            move |_| {
                                let category = active_category.read().clone();
                                let picked = pick_files_for_type(&category);
                                if !picked.is_empty() {
                                    let mut new_mgr = resources.clone();
                                    for (name, path) in picked {
                                        let item = ResourceItem::new_file(name, &path, category.clone());
                                        new_mgr.resources.push(item);
                                    }
                                    on_change.call(new_mgr);
                                }
                            }
                        },
                        "Add {cat}..."
                    }
                }
            }

            // Category tabs (like VS resource editor)
            div {
                style: "display: flex; background: #fafafa; border-bottom: 1px solid #ddd; gap: 0;",
                button {
                    style: if *active_category.read() == ResourceType::String { tab_active_style } else { tab_style },
                    onclick: move |_| active_category.set(ResourceType::String),
                    "Strings ({count_strings})"
                }
                button {
                    style: if *active_category.read() == ResourceType::Image { tab_active_style } else { tab_style },
                    onclick: move |_| active_category.set(ResourceType::Image),
                    "Images ({count_images})"
                }
                button {
                    style: if *active_category.read() == ResourceType::Icon { tab_active_style } else { tab_style },
                    onclick: move |_| active_category.set(ResourceType::Icon),
                    "Icons ({count_icons})"
                }
                button {
                    style: if *active_category.read() == ResourceType::Audio { tab_active_style } else { tab_style },
                    onclick: move |_| active_category.set(ResourceType::Audio),
                    "Audio ({count_audio})"
                }
                button {
                    style: if *active_category.read() == ResourceType::File { tab_active_style } else { tab_style },
                    onclick: move |_| active_category.set(ResourceType::File),
                    "Files ({count_files})"
                }
                button {
                    style: if *active_category.read() == ResourceType::Other { tab_active_style } else { tab_style },
                    onclick: move |_| active_category.set(ResourceType::Other),
                    "Other ({count_other})"
                }
            }

            // Table
            div {
                style: "flex: 1; overflow-y: auto; padding: 10px;",
                table {
                    style: "width: 100%; border-collapse: collapse;",
                    thead {
                        tr {
                            th { style: "text-align: left; border-bottom: 2px solid #ddd; padding: 8px;", "Name" }
                            th { style: "text-align: left; border-bottom: 2px solid #ddd; padding: 8px;",
                                {match *active_category.read() {
                                    ResourceType::String => "Value",
                                    ResourceType::Image | ResourceType::Icon | ResourceType::Audio | ResourceType::File => "File Path",
                                    ResourceType::Other => "Data",
                                }}
                            }
                            th { style: "text-align: left; border-bottom: 2px solid #ddd; padding: 8px;", "Comment" }
                            th { style: "text-align: left; border-bottom: 2px solid #ddd; padding: 8px; width: 120px;", "Actions" }
                        }
                    }
                    tbody {
                        if filtered.is_empty() {
                            tr {
                                td {
                                    colspan: "4",
                                    style: "padding: 30px; text-align: center; color: #999; font-style: italic;",
                                    if is_file_cat {
                                        "No {active_category.read()} resources. Click \"Add {active_category.read()}...\" to browse for files."
                                    } else {
                                        "No {active_category.read()} resources. Add one below."
                                    }
                                }
                            }
                        }
                        for (idx, res) in filtered.iter() {
                            {
                                let real_idx = *idx;
                                let res_name = res.name.clone();
                                let res_val = res.value.clone();
                                let res_comment = res.comment.clone().unwrap_or_default();
                                let res_type = res.resource_type.clone();
                                let is_file_row = is_file_category(&res_type);

                                let on_change = props.on_change;
                                let resources = props.resources.clone();

                                rsx! {
                                    tr {
                                        key: "{real_idx}",
                                        // Name column — editable
                                        td { 
                                            style: "padding: 4px; border-bottom: 1px solid #eee;",
                                            input {
                                                style: "width: 100%; border: 1px solid transparent; background: transparent; padding: 2px 4px;",
                                                value: "{res_name}",
                                                oninput: {
                                                    let resources = resources.clone();
                                                    move |evt| {
                                                        let mut new_mgr = resources.clone();
                                                        if let Some(r) = new_mgr.resources.get_mut(real_idx) {
                                                            r.name = evt.value().clone();
                                                        }
                                                        on_change.call(new_mgr);
                                                    }
                                                }
                                            }
                                        }
                                        // Value/path column
                                        td { 
                                            style: "padding: 4px; border-bottom: 1px solid #eee;",
                                            if is_file_row {
                                                // File path display + browse button
                                                div {
                                                    style: "display: flex; align-items: center; gap: 4px;",
                                                    span {
                                                        style: "flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12px; color: #555; padding: 2px 4px;",
                                                        title: "{res_val}",
                                                        {
                                                            // Show just the filename for readability
                                                            Path::new(&res_val)
                                                                .file_name()
                                                                .and_then(|f| f.to_str())
                                                                .unwrap_or(&res_val)
                                                                .to_string()
                                                        }
                                                    }
                                                    button {
                                                        style: "padding: 1px 8px; border: 1px solid #ccc; background: #f0f0f0; border-radius: 3px; cursor: pointer; font-size: 11px; white-space: nowrap;",
                                                        onclick: {
                                                            let resources = resources.clone();
                                                            let res_type = res_type.clone();
                                                            move |_| {
                                                                if let Some((_name, path)) = pick_single_file_for_type(&res_type) {
                                                                    let mut new_mgr = resources.clone();
                                                                    if let Some(r) = new_mgr.resources.get_mut(real_idx) {
                                                                        r.value = path.clone();
                                                                        r.file_name = Some(path);
                                                                    }
                                                                    on_change.call(new_mgr);
                                                                }
                                                            }
                                                        },
                                                        "Browse..."
                                                    }
                                                }
                                            } else {
                                                // Editable text for strings/other
                                                input {
                                                    style: "width: 100%; border: 1px solid transparent; background: transparent; padding: 2px 4px;",
                                                    value: "{res_val}",
                                                    oninput: {
                                                        let resources = resources.clone();
                                                        move |evt| {
                                                            let mut new_mgr = resources.clone();
                                                            if let Some(r) = new_mgr.resources.get_mut(real_idx) {
                                                                r.value = evt.value().clone();
                                                            }
                                                            on_change.call(new_mgr);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        // Comment column — always editable
                                        td { 
                                            style: "padding: 4px; border-bottom: 1px solid #eee;",
                                            input {
                                                style: "width: 100%; border: 1px solid transparent; background: transparent; padding: 2px 4px;",
                                                value: "{res_comment}",
                                                oninput: {
                                                    let resources = resources.clone();
                                                    move |evt| {
                                                        let mut new_mgr = resources.clone();
                                                        if let Some(r) = new_mgr.resources.get_mut(real_idx) {
                                                            r.comment = Some(evt.value().clone());
                                                        }
                                                        on_change.call(new_mgr);
                                                    }
                                                }
                                            }
                                        }
                                        // Actions column
                                        td {
                                            style: "padding: 4px; border-bottom: 1px solid #eee; white-space: nowrap;",
                                            button {
                                                style: "color: red; cursor: pointer; border: none; background: none; font-size: 12px;",
                                                onclick: {
                                                    let resources = resources.clone();
                                                    move |_| {
                                                        let mut new_mgr = resources.clone();
                                                        new_mgr.resources.remove(real_idx);
                                                        on_change.call(new_mgr);
                                                    }
                                                },
                                                "Remove"
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Footer — only for string/other categories (file categories use the top Add button)
            if !is_file_cat {
                div {
                    style: "padding: 10px; border-top: 1px solid #ccc; display: flex; gap: 10px; align-items: center; background: #f9f9f9;",
                    input {
                        style: "flex: 1; padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px;",
                        placeholder: "Name",
                        value: "{new_res_name}",
                        oninput: move |evt| new_res_name.set(evt.value().clone())
                    }
                    input {
                        style: "flex: 2; padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px;",
                        placeholder: {match *active_category.read() {
                            ResourceType::String => "Value",
                            _ => "Data",
                        }},
                        value: "{new_res_value}",
                        oninput: move |evt| new_res_value.set(evt.value().clone())
                    }
                    input {
                        style: "flex: 1; padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px;",
                        placeholder: "Comment (optional)",
                        value: "{new_res_comment}",
                        oninput: move |evt| new_res_comment.set(evt.value().clone())
                    }
                    button {
                        style: "padding: 4px 12px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer; white-space: nowrap;",
                        onclick: move |_| {
                            if !new_res_name.read().is_empty() {
                                let category = active_category.read().clone();
                                let mut new_mgr = props.resources.clone();
                                let item = ResourceItem::new_string(
                                    new_res_name.read().clone(),
                                    new_res_value.read().clone(),
                                );
                                let mut item = item;
                                item.resource_type = category;
                                item.comment = if new_res_comment.read().is_empty() { None } else { Some(new_res_comment.read().clone()) };
                                new_mgr.resources.push(item);
                                props.on_change.call(new_mgr);
                                new_res_name.set(String::new());
                                new_res_value.set(String::new());
                                new_res_comment.set(String::new());
                            }
                        },
                        "Add {active_category.read()}"
                    }
                }
            }
        }
    }
}
