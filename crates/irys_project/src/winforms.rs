use crate::errors::{SaveError, SaveResult};
use crate::project::{FormFormat, FormModule};
use irys_forms::{EventBinding, EventType, Form};
use std::fs;
use std::path::Path;

fn event_type_from_name(name: &str) -> Option<EventType> {
    match name.to_lowercase().as_str() {
        "click" => Some(EventType::Click),
        "dblclick" => Some(EventType::DblClick),
        "load" => Some(EventType::Load),
        "unload" => Some(EventType::Unload),
        "change" => Some(EventType::Change),
        "keypress" => Some(EventType::KeyPress),
        "keydown" => Some(EventType::KeyDown),
        "keyup" => Some(EventType::KeyUp),
        "mousedown" => Some(EventType::MouseDown),
        "mouseup" => Some(EventType::MouseUp),
        "mousemove" => Some(EventType::MouseMove),
        "gotfocus" => Some(EventType::GotFocus),
        "lostfocus" => Some(EventType::LostFocus),
        _ => None,
    }
}

fn apply_vbnet_handles(form: &mut Form, cls: &irys_parser::ClassDecl) {
    for method in &cls.methods {
        if let irys_parser::ast::decl::MethodDecl::Sub(sub) = method {
            let Some(handles) = sub.handles.as_ref() else { continue };

            for handle in handles {
                let parts: Vec<&str> = handle.split('.').collect();
                if parts.len() < 2 {
                    continue;
                }

                let event_part = parts.last().unwrap();
                let control_part = parts.get(parts.len() - 2).unwrap();

                let control_name = if control_part.eq_ignore_ascii_case("me")
                    || control_part.eq_ignore_ascii_case("mybase")
                    || control_part.eq_ignore_ascii_case(&form.name)
                {
                    form.name.clone()
                } else {
                    control_part.to_string()
                };

                let Some(event_type) = event_type_from_name(event_part) else { continue };

                if !control_name.eq_ignore_ascii_case(&form.name)
                    && form.get_control_by_name(&control_name).is_none()
                    && !matches!(event_type, EventType::Load | EventType::Unload)
                {
                    continue;
                }

                let handler_name = sub.name.as_str().to_string();
                let already = form.event_bindings.iter().any(|b| {
                    b.control_name.eq_ignore_ascii_case(&control_name)
                        && b.event_type == event_type
                        && b.handler_name.eq_ignore_ascii_case(&handler_name)
                });

                if !already {
                    form.add_event_binding(EventBinding::with_handler(
                        control_name.clone(),
                        event_type.clone(),
                        handler_name,
                    ));
                }
            }
        }
    }
}

pub fn load_form_vb(form_path: &Path) -> SaveResult<FormModule> {
    let user_code = fs::read_to_string(form_path)?;

    let stem = form_path.file_stem().unwrap_or_default().to_string_lossy().to_string();
    let parent = form_path.parent().unwrap_or(Path::new("."));
    let designer_path = parent.join(format!("{}.Designer.vb", stem));

    let designer_code = if designer_path.exists() {
        fs::read_to_string(&designer_path)?
    } else {
        String::new()
    };

    let combined = format!("{}\n{}", designer_code, user_code);
    eprintln!("[DEBUG] load_form_vb: designer_path={:?} exists={}", designer_path, designer_path.exists());
    eprintln!("[DEBUG] combined code length: {} chars", combined.len());
    let program = irys_parser::parse_program(&combined)
        .map_err(|e| { eprintln!("[DEBUG] PARSE ERROR: {}", e); SaveError::Parse(format!("{}", e)) })?;
    eprintln!("[DEBUG] Parsed OK: {} declarations", program.declarations.len());

    let mut merged_class: Option<irys_parser::ClassDecl> = None;
    for decl in &program.declarations {
        if let irys_parser::Declaration::Class(cls) = decl {
            if let Some(ref mut existing) = merged_class {
                existing.fields.extend(cls.fields.clone());
                existing.methods.extend(cls.methods.clone());
                existing.properties.extend(cls.properties.clone());
                if existing.inherits.is_none() {
                    existing.inherits = cls.inherits.clone();
                }
            } else {
                merged_class = Some(cls.clone());
            }
        }
    }

    let mut form = if let Some(cls) = &merged_class {
        eprintln!("[DEBUG] merged_class: '{}' with {} methods, {} fields", cls.name.as_str(), cls.methods.len(), cls.fields.len());
        let extracted = irys_forms::serialization::designer_parser::extract_form_from_designer(cls);
        eprintln!("[DEBUG] extract_form_from_designer returned: {}", if extracted.is_some() { "Some" } else { "None" });
        match extracted {
            Some(f) => {
                eprintln!("[DEBUG] Extracted form '{}' with {} controls, size={}x{}", f.name, f.controls.len(), f.width, f.height);
                f
            }
            None => {
                eprintln!("[DEBUG] extract_form_from_designer returned None, creating empty form");
                Form::new(cls.name.as_str())
            }
        }
    } else {
        eprintln!("[DEBUG] No merged_class found, creating empty form '{}'", stem);
        Form::new(&stem)
    };

    if let Some(cls) = &merged_class {
        apply_vbnet_handles(&mut form, cls);
    }

    Ok(FormModule::new_vbnet(form, designer_code, user_code))
}

pub fn save_form_vb(form_module: &FormModule, dir: &Path) -> SaveResult<()> {
    let name = &form_module.form.name;

    if let FormFormat::VbNet { designer_code, user_code } = &form_module.format {
        let designer_path = dir.join(format!("{}.Designer.vb", name));
        fs::write(&designer_path, designer_code)?;

        let user_path = dir.join(format!("{}.vb", name));
        fs::write(&user_path, user_code)?;
    }

    Ok(())
}
