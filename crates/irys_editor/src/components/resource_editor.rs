
use dioxus::prelude::*;
use irys_project::ResourceManager;
use irys_project::ResourceItem;

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

    rsx! {
        div {
            class: "resource-editor",
            style: "display: flex; flex-direction: column; height: 100%; border: 1px solid #ccc; background: white;",
            
            // Header
            div {
                style: "padding: 10px; background: #f0f0f0; border-bottom: 1px solid #ccc; font-weight: bold;",
                "Project Resources (.resx)"
            }

            // Table
            div {
                style: "flex: 1; overflow-y: auto; padding: 10px;",
                table {
                    style: "width: 100%; border-collapse: collapse;",
                    thead {
                        tr {
                            th { style: "text-align: left; border-bottom: 2px solid #ddd; padding: 8px;", "Name" }
                            th { style: "text-align: left; border-bottom: 2px solid #ddd; padding: 8px;", "Value" }
                            th { style: "text-align: left; border-bottom: 2px solid #ddd; padding: 8px;", "Comment" }
                            th { style: "text-align: left; border-bottom: 2px solid #ddd; padding: 8px;", "Actions" }
                        }
                    }
                    tbody {
                        for (idx, res) in props.resources.resources.iter().enumerate() {
                            {
                                let res_name = res.name.clone();
                                let res_val = res.value.clone();
                                let res_comment = res.comment.clone().unwrap_or_default();
                                
                                // Clone once for all closures in this row if possible, or per closure
                                let on_change = props.on_change;
                                let resources = props.resources.clone();

                                rsx! {
                                    tr {
                                        key: "{idx}",
                                        td { 
                                            style: "padding: 4px; border-bottom: 1px solid #eee;",
                                            input {
                                                style: "width: 100%; border: 1px solid transparent; background: transparent;",
                                                value: "{res_name}",
                                                oninput: {
                                                    let resources = resources.clone();
                                                    move |evt| {
                                                        let mut new_mgr = resources.clone();
                                                        if let Some(r) = new_mgr.resources.get_mut(idx) {
                                                            r.name = evt.value().clone();
                                                        }
                                                        on_change.call(new_mgr);
                                                    }
                                                }
                                            }
                                        }
                                        td { 
                                            style: "padding: 4px; border-bottom: 1px solid #eee;",
                                            input {
                                                style: "width: 100%; border: 1px solid transparent; background: transparent;",
                                                value: "{res_val}",
                                                oninput: {
                                                    let resources = resources.clone();
                                                    move |evt| {
                                                        let mut new_mgr = resources.clone();
                                                        if let Some(r) = new_mgr.resources.get_mut(idx) {
                                                            r.value = evt.value().clone();
                                                        }
                                                        on_change.call(new_mgr);
                                                    }
                                                }
                                            }
                                        }
                                        td { 
                                            style: "padding: 4px; border-bottom: 1px solid #eee;",
                                            input {
                                                style: "width: 100%; border: 1px solid transparent; background: transparent;",
                                                value: "{res_comment}",
                                                oninput: {
                                                    let resources = resources.clone();
                                                    move |evt| {
                                                        let mut new_mgr = resources.clone();
                                                        if let Some(r) = new_mgr.resources.get_mut(idx) {
                                                            r.comment = Some(evt.value().clone());
                                                        }
                                                        on_change.call(new_mgr);
                                                    }
                                                }
                                            }
                                        }
                                        td {
                                            style: "padding: 4px; border-bottom: 1px solid #eee;",
                                            button {
                                                style: "color: red; cursor: pointer; border: none; background: none;",
                                                onclick: {
                                                    let resources = resources.clone();
                                                    move |_| {
                                                        let mut new_mgr = resources.clone();
                                                        new_mgr.resources.remove(idx);
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

            // Footer (Add New)
            div {
                style: "padding: 10px; border-top: 1px solid #ccc; display: flex; gap: 10px; background: #f9f9f9;",
                input {
                    placeholder: "Name",
                    value: "{new_res_name}",
                    oninput: move |evt| new_res_name.set(evt.value().clone())
                }
                input {
                    placeholder: "Value",
                    value: "{new_res_value}",
                    oninput: move |evt| new_res_value.set(evt.value().clone())
                }
                input {
                    placeholder: "Comment (optional)",
                    value: "{new_res_comment}",
                    oninput: move |evt| new_res_comment.set(evt.value().clone())
                }
                button {
                    style: "padding: 4px 12px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer;",
                    onclick: move |_| {
                        if !new_res_name.read().is_empty() {
                            let mut new_mgr = props.resources.clone();
                            new_mgr.resources.push(ResourceItem {
                                name: new_res_name.read().clone(),
                                value: new_res_value.read().clone(),
                                comment: if new_res_comment.read().is_empty() { None } else { Some(new_res_comment.read().clone()) },
                            });
                            props.on_change.call(new_mgr);
                            new_res_name.set(String::new());
                            new_res_value.set(String::new());
                            new_res_comment.set(String::new());
                        }
                    },
                    "Add Resource"
                }
            }
        }
    }
}
