//! PHP process environment functions whose names do not map to WASI imports.
use std::path::PathBuf;
use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Framework, Value, heap};

const CWD_KEY: &str = "__vybe_php_virtual_cwd";

fn map_string(value: &Value, key: &str) -> Option<String> {
    let Value::Object(object) = value else {
        return None;
    };
    let locked = object.lock().unwrap();
    let ObjectKind::Map(entries) = &locked.kind else {
        return None;
    };
    match entries.get(&Value::String(Arc::from(key))) {
        Some(Value::String(value)) => Some(value.to_string()),
        _ => None,
    }
}

fn current_php_dir(ctx: &vybe_runtime::HostContext<'_>) -> PathBuf {
    if let Some(path) = map_string(&ctx.get_global("__php_env_overlay"), CWD_KEY) {
        return PathBuf::from(path);
    }
    if let Some(script) = map_string(&ctx.get_global("__wasi_http_server_env"), "SCRIPT_FILENAME") {
        if let Some(parent) = PathBuf::from(script).parent() {
            return parent.to_path_buf();
        }
    }
    std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."))
}

fn callable_reflection_name(value: &Value) -> Option<String> {
    match value {
        Value::String(name) => Some(name.rsplit("::").next().unwrap_or(name).to_string()),
        Value::Object(object) => {
            let object = object.lock().unwrap();
            let method = match &object.kind {
                ObjectKind::Array(items) => items.get(1),
                ObjectKind::Map(items) => items.get(&Value::I32(1)),
                _ => object.properties.get("1"),
            };
            if let Some(Value::String(name)) = method {
                return Some(name.to_string());
            }
            if let ObjectKind::Function(function) = &object.kind {
                return Some(
                    function
                        .name
                        .as_deref()
                        .filter(|name| !name.starts_with('<'))
                        .unwrap_or("{closure}")
                        .to_string(),
                );
            }
            None
        }
        _ => None,
    }
}

fn callable_pair(value: &Value) -> Option<(Value, Value)> {
    let Value::Object(object) = value else {
        return None;
    };
    let object = object.lock().unwrap();
    match &object.kind {
        ObjectKind::Array(items) if items.len() >= 2 => Some((items[0].clone(), items[1].clone())),
        ObjectKind::Map(items) => Some((
            items.get(&Value::I32(0))?.clone(),
            items.get(&Value::I32(1))?.clone(),
        )),
        _ => Some((
            object.properties.get("0")?.clone(),
            object.properties.get("1")?.clone(),
        )),
    }
}

fn closure_callable(closure: &Value) -> Option<Value> {
    let Value::Object(object) = closure else {
        return None;
    };
    object
        .lock()
        .unwrap()
        .properties
        .get("__php_reflection_target")
        .cloned()
}

fn collect_callable_properties(value: &Value, names: &mut Vec<String>) {
    let Value::Object(object) = value else {
        return;
    };
    let object = object.lock().unwrap();
    for (name, member) in &object.properties {
        if name.starts_with("__") || name == "constructor" {
            continue;
        }
        if let Value::Object(function) = member {
            if matches!(function.lock().unwrap().kind, ObjectKind::Function(_))
                && !names.iter().any(|seen| seen == name)
            {
                names.push(name.clone());
            }
        }
    }
}

fn header_name(server_key: &str) -> Option<String> {
    let raw = match server_key {
        "CONTENT_TYPE" => return Some("Content-Type".to_string()),
        "CONTENT_LENGTH" => return Some("Content-Length".to_string()),
        "CONTENT_MD5" => return Some("Content-Md5".to_string()),
        _ => server_key.strip_prefix("HTTP_")?,
    };
    Some(
        raw.split('_')
            .map(|part| {
                let mut characters = part.chars();
                match characters.next() {
                    Some(first) => format!(
                        "{}{}",
                        first.to_ascii_uppercase(),
                        characters.as_str().to_ascii_lowercase()
                    ),
                    None => String::new(),
                }
            })
            .collect::<Vec<_>>()
            .join("-"),
    )
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:environment",
        "getCwd",
        Box::new(|ctx, _| {
            Value::String(Arc::from(current_php_dir(ctx).to_string_lossy().as_ref()))
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "chdir",
        Box::new(|ctx, args| {
            let Some(Value::String(path)) = args.first() else {
                return Value::Bool(false);
            };
            let requested = PathBuf::from(path.as_ref());
            let absolute = if requested.is_absolute() {
                requested
            } else {
                current_php_dir(ctx).join(requested)
            };
            let Ok(absolute) = absolute.canonicalize() else {
                return Value::Bool(false);
            };
            if !absolute.is_dir() {
                return Value::Bool(false);
            }
            let overlay = match ctx.get_global("__php_env_overlay") {
                Value::Object(object) => object,
                _ => {
                    let mut map = Object::new();
                    map.kind = ObjectKind::Map(Default::default());
                    let object = heap::alloc(map);
                    ctx.set_global("__php_env_overlay", Value::Object(object.clone()));
                    object
                }
            };
            if let ObjectKind::Map(entries) = &mut overlay.lock().unwrap().kind {
                entries.insert(
                    Value::String(Arc::from(CWD_KEY)),
                    Value::String(Arc::from(absolute.to_string_lossy().as_ref())),
                );
            }
            Value::Bool(true)
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "attachCallableReflection",
        Box::new(|_, args| {
            let closure = args.first().cloned().unwrap_or(Value::Null);
            if let (Value::Object(object), Some(callable)) = (&closure, args.get(1)) {
                object
                    .lock()
                    .unwrap()
                    .properties
                    .insert("__php_reflection_target".to_string(), callable.clone());
                if let Some(name) = callable_reflection_name(callable) {
                    object.lock().unwrap().properties.insert(
                        "__php_reflection_name".to_string(),
                        Value::String(name.into()),
                    );
                }
            }
            closure
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "reflectionFunctionName",
        Box::new(|_, args| {
            if let Some(Value::Object(object)) = args.first() {
                if let Some(Value::String(name)) = object
                    .lock()
                    .unwrap()
                    .properties
                    .get("__php_reflection_name")
                {
                    return Value::String(name.clone());
                }
            }
            Value::String(
                callable_reflection_name(args.first().unwrap_or(&Value::Null))
                    .unwrap_or_else(|| "{closure}".to_string())
                    .into(),
            )
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "closureThis",
        Box::new(|_, args| {
            closure_callable(args.first().unwrap_or(&Value::Null))
                .as_ref()
                .and_then(callable_pair)
                .map(|(receiver, _)| {
                    if matches!(receiver, Value::Object(_)) {
                        receiver
                    } else {
                        Value::Null
                    }
                })
                .unwrap_or(Value::Null)
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "closureCalledClass",
        Box::new(|_, args| {
            let class = closure_callable(args.first().unwrap_or(&Value::Null))
                .as_ref()
                .and_then(callable_pair)
                .and_then(|(receiver, _)| match receiver {
                    Value::String(name) => Some(name.to_string()),
                    Value::Object(object) => {
                        object.lock().unwrap().properties.get("__type").and_then(
                            |value| match value {
                                Value::String(name) => Some(name.to_string()),
                                _ => None,
                            },
                        )
                    }
                    _ => None,
                });
            match class {
                Some(class) => {
                    let mut reflected = Object::new();
                    reflected
                        .properties
                        .insert("name".to_string(), Value::String(class.into()));
                    Value::Object(heap::alloc(reflected))
                }
                None => Value::Null,
            }
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "reflectionClassName",
        Box::new(|_, args| match args.first() {
            Some(Value::String(name)) => Value::String(name.clone()),
            Some(Value::Object(object)) => object
                .lock()
                .unwrap()
                .properties
                .get("__type")
                .and_then(|value| match value {
                    Value::String(name) => Some(Value::String(name.clone())),
                    _ => None,
                })
                .unwrap_or(Value::String("stdClass".into())),
            _ => Value::String("".into()),
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "getClassMethods",
        Box::new(|ctx, args| {
            let mut names = Vec::new();
            if let Some(value) = args.first() {
                collect_callable_properties(value, &mut names);
                if let Value::String(class_name) = value {
                    let class = ctx.get_global(&class_name.replace('\\', "."));
                    collect_callable_properties(&class, &mut names);
                    if let Value::Object(class) = class {
                        if let Some(prototype) =
                            class.lock().unwrap().properties.get("prototype").cloned()
                        {
                            collect_callable_properties(&prototype, &mut names);
                        }
                    }
                }
            }
            Value::Object(heap::alloc(Object::new_array(
                names
                    .into_iter()
                    .map(|name| Value::String(name.into()))
                    .collect(),
            )))
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "attachClosureReflection",
        Box::new(|_, args| {
            let closure = args.first().cloned().unwrap_or(Value::Null);
            if let (Value::Object(object), Some(params)) = (&closure, args.get(1)) {
                object
                    .lock()
                    .unwrap()
                    .properties
                    .insert("__php_reflection_params".to_string(), params.clone());
            }
            closure
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "closureReflectionParams",
        Box::new(|_, args| {
            if let Some(Value::Object(object)) = args.first() {
                if let Some(params) = object
                    .lock()
                    .unwrap()
                    .properties
                    .get("__php_reflection_params")
                {
                    return params.clone();
                }
            }
            let mut empty = Object::new();
            empty.kind = ObjectKind::Array(Vec::new());
            Value::Object(heap::alloc(empty))
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "isClosure",
        Box::new(|_, args| {
            let is_closure = matches!(args.first(), Some(Value::Object(object))
            if matches!(object.lock().unwrap().kind, ObjectKind::Function(_)));
            Value::Bool(is_closure)
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "getallheaders",
        Box::new(|ctx, _| {
            let mut headers = Object::new();
            headers.kind = ObjectKind::Map(Default::default());
            if let Value::Object(server) = ctx.get_global("$_SERVER") {
                let server = server.lock().unwrap();
                if let (ObjectKind::Map(source), ObjectKind::Map(destination)) =
                    (&server.kind, &mut headers.kind)
                {
                    for (key, value) in source {
                        // The HTTP server leaves CONTENT_TYPE and CONTENT_LENGTH
                        // null when the client did not send those headers. PHP's
                        // getallheaders() lists only headers that were present.
                        if matches!(value, Value::Null | Value::Undefined) {
                            continue;
                        }
                        if let Value::String(key) = key {
                            if let Some(name) = header_name(key) {
                                destination.insert(Value::String(name.into()), value.clone());
                            }
                        }
                    }
                }
            }
            Value::Object(heap::alloc(headers))
        }),
    );
    fw.register_host_fn(
        "php:environment",
        "gethostname",
        Box::new(|_, _| {
            #[cfg(unix)]
            {
                let mut bytes = [0u8; 256];
                let status = unsafe { libc::gethostname(bytes.as_mut_ptr().cast(), bytes.len()) };
                if status == 0 {
                    let end = bytes
                        .iter()
                        .position(|byte| *byte == 0)
                        .unwrap_or(bytes.len());
                    return Value::String(
                        String::from_utf8_lossy(&bytes[..end]).into_owned().into(),
                    );
                }
            }
            Value::Bool(false)
        }),
    );
}
