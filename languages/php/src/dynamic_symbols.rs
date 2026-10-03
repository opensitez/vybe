//! Runtime symbol lookup for PHP names that become available after a file was compiled.
use std::sync::OnceLock;
use vybe_runtime::value::ObjectKind;
use vybe_runtime::{Framework, Value};

fn class_object(
    ctx: &mut vybe_runtime::HostContext,
    name: &str,
) -> Option<std::sync::Arc<std::sync::Mutex<vybe_runtime::value::Object>>> {
    let bare = name.trim_start_matches('\\');
    for candidate in [
        bare.to_string(),
        bare.replace('\\', "."),
        format!("\\{bare}"),
    ] {
        if let Value::Object(object) = ctx.get_global(&candidate) {
            return Some(object);
        }
    }
    None
}

fn named_field(
    ctx: &vybe_runtime::HostContext<'_>,
    object: &vybe_runtime::value::Object,
    name: &str,
) -> Option<Value> {
    if let Some(value) = object.properties.get(name) {
        return Some(value.clone());
    }
    ctx.declared_fields(object.type_id)
        .iter()
        .position(|(field, _)| field == name)
        .and_then(|index| object.fields.get(index).cloned())
}

fn has_method_on_chain(
    ctx: &vybe_runtime::HostContext<'_>,
    object: &std::sync::Arc<std::sync::Mutex<vybe_runtime::value::Object>>,
    method: &str,
) -> bool {
    let mut current = Some(object.clone());
    while let Some(obj) = current {
        let locked = obj.lock().unwrap();
        let fields = ctx.declared_fields(locked.type_id);
        if locked.properties.iter().any(|(key, value)| key.eq_ignore_ascii_case(method)
            && matches!(value, Value::Object(function)
                if matches!(function.lock().unwrap().kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_))))
            || fields.iter().enumerate().any(|(index, (key, _))| key.eq_ignore_ascii_case(method)
                && matches!(locked.fields.get(index), Some(Value::Object(function))
                    if matches!(function.lock().unwrap().kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_)))) {
            return true;
        }
        current = match named_field(ctx, &locked, "__proto__") {
            Some(Value::Object(parent)) => Some(parent.clone()),
            _ => None,
        };
    }
    false
}

fn has_instance_signature(
    ctx: &vybe_runtime::HostContext<'_>,
    object: &std::sync::Arc<std::sync::Mutex<vybe_runtime::value::Object>>,
    method: &str,
) -> bool {
    let prefix = format!("{}$sig", method.to_ascii_lowercase());
    let locked = object.lock().unwrap();
    let fields = ctx.declared_fields(locked.type_id);
    locked
        .properties
        .keys()
        .chain(fields.iter().map(|(name, _)| name))
        .any(|name| name.to_ascii_lowercase().starts_with(&prefix))
}

fn class_method_kind(
    ctx: &mut vybe_runtime::HostContext,
    name: &str,
    method: &str,
) -> (bool, bool) {
    let Some(class) = class_object(ctx, name) else {
        return ctx
            .class_method_kind(name, method)
            .unwrap_or((false, false));
    };
    let prototype = {
        let locked = class.lock().unwrap();
        match named_field(ctx, &locked, "prototype") {
            Some(Value::Object(prototype)) => Some(prototype.clone()),
            _ => None,
        }
    };
    let instance = has_instance_signature(ctx, &class, method)
        || prototype
            .as_ref()
            .is_some_and(|prototype| has_method_on_chain(ctx, prototype, method));
    let static_method = !instance && has_method_on_chain(ctx, &class, method);
    (instance, static_method)
}

fn profile_builtin_exists(name: &str) -> bool {
    static PROFILE: OnceLock<vybe_runtime::profile::LanguageProfile> = OnceLock::new();
    let name = name.trim_start_matches('\u{5c}').to_ascii_lowercase();
    if name.contains('\u{5c}') || name.contains('.') {
        return false;
    }
    PROFILE
        .get_or_init(|| {
            vybe_runtime::profile::parse_profile(crate::profile_source())
                .expect("PHP profile must parse")
        })
        .lookup_builtin(&name)
        .is_some()
        || name.starts_with("sodium_")
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:dynamic",
        "methodExists",
        Box::new(|ctx, args| {
            let (Some(Value::String(class)), Some(Value::String(method))) =
                (args.first(), args.get(1))
            else {
                return Value::Bool(false);
            };
            let (instance, static_method) = class_method_kind(ctx, class, method);
            Value::Bool(instance || static_method)
        }),
    );
    fw.register_host_fn(
        "php:dynamic",
        "methodIsStatic",
        Box::new(|ctx, args| {
            let (Some(Value::String(class)), Some(Value::String(method))) =
                (args.first(), args.get(1))
            else {
                return Value::Bool(false);
            };
            let (instance, static_method) = class_method_kind(ctx, class, method);
            Value::Bool(!instance && static_method)
        }),
    );
    fw.register_host_fn(
        "php:dynamic",
        "methodReflectionParams",
        Box::new(|ctx, args| {
            let (Some(Value::String(class)), Some(Value::String(method))) =
                (args.first(), args.get(1))
            else {
                return Value::Undefined;
            };
            let mut current = class_object(ctx, class);
            let key = Value::String(method.to_ascii_lowercase().into());
            while let Some(object) = current {
                let locked = object.lock().unwrap();
                if let Some(Value::Object(params)) =
                    named_field(ctx, &locked, "__php_reflection_method_params")
                {
                    let params = params.lock().unwrap();
                    if let Some(value) = params.properties.get(method.to_ascii_lowercase().as_str())
                    {
                        return value.clone();
                    }
                    if let ObjectKind::Map(entries) = &params.kind {
                        if let Some(value) = entries.get(&key) {
                            return value.clone();
                        }
                    }
                }
                current = match named_field(ctx, &locked, "__proto__") {
                    Some(Value::Object(parent)) => Some(parent),
                    _ => None,
                };
            }
            Value::Undefined
        }),
    );
    fw.register_host_fn(
        "php:fiber",
        "isFiber",
        Box::new(|_, args| {
            Value::Bool(matches!(args.first(), Some(Value::Object(object))
            if matches!(object.lock().unwrap().kind, ObjectKind::Continuation(_))))
        }),
    );
    fw.register_host_fn(
        "php:dynamic",
        "builtinExists",
        Box::new(|_, args| {
            Value::Bool(
                matches!(args.first(), Some(Value::String(name)) if profile_builtin_exists(name)),
            )
        }),
    );
    fw.register_host_fn(
        "vybe:php",
        "include_scope",
        Box::new(|_, _| vybe_compiler::dynamic::php_include_scope()),
    );
    fw.register_host_fn(
        "php:dynamic",
        "functionByName",
        Box::new(|ctx, args| {
            let Some(Value::String(name)) = args.first() else {
                return Value::Undefined;
            };
            let canonical = name
                .split(['\\', '.'])
                .filter(|part| !part.is_empty())
                .collect::<Vec<_>>()
                .join(".")
                .to_ascii_lowercase();
            ctx.get_global(&format!("__vybe_func${canonical}"))
        }),
    );
    fw.register_host_fn(
        "php:dynamic",
        "globalByName",
        Box::new(|ctx, args| {
            let Some(Value::String(name)) = args.first() else {
                return Value::Undefined;
            };
            let direct = ctx.get_global(name);
            if !matches!(direct, Value::Undefined) {
                return direct;
            }
            let normalized = name.trim_start_matches('\\').replace('\\', ".");
            if normalized == name.as_ref() {
                Value::Undefined
            } else {
                ctx.get_global(&normalized)
            }
        }),
    );
    fw.register_host_fn(
        "php:dynamic",
        "autoloadName",
        Box::new(|_, args| {
            let Some(Value::String(name)) = args.first() else {
                return Value::Undefined;
            };
            Value::String(std::sync::Arc::from(name.trim_start_matches('\\')))
        }),
    );
    fw.register_host_fn(
        "php:dynamic",
        "constructNamed",
        Box::new(|ctx, args| {
            let (
                Some(constructor @ Value::Object(function)),
                Some(Value::Object(values)),
                Some(Value::Object(names)),
            ) = (args.first(), args.get(1), args.get(2))
            else {
                return Value::Undefined;
            };
            let class_name = match &function.lock().unwrap().kind {
                ObjectKind::Function(function) => function.name.clone(),
                _ => None,
            };
            let Some(class_name) = class_name else {
                return Value::Undefined;
            };
            let helper_name = format!("__{class_name}_ctor_0");
            let Some(mut parameters) = ctx.source_parameter_names_for_chunk(&helper_name) else {
                return Value::Undefined;
            };
            while parameters.last().is_some_and(Option::is_none) {
                parameters.pop();
            }
            let values = match &values.lock().unwrap().kind {
                ObjectKind::Array(values) => values.clone(),
                _ => return Value::Undefined,
            };
            let names = match &names.lock().unwrap().kind {
                ObjectKind::Array(names) => names.clone(),
                _ => return Value::Undefined,
            };
            if values.len() != names.len() {
                return Value::Undefined;
            }

            let mut ordered = vec![None; parameters.len()];
            let mut extra = Vec::new();
            let mut next_positional = 0usize;
            for (name, value) in names.iter().zip(values) {
                match name {
                    Value::String(name) => {
                        let Some(index) = parameters.iter().position(|parameter| {
                            parameter.as_deref().is_some_and(|parameter| {
                                parameter.trim_start_matches('$') == name.as_ref()
                            })
                        }) else {
                            return Value::Undefined;
                        };
                        if ordered[index].replace(value).is_some() {
                            return Value::Undefined;
                        }
                    }
                    Value::Undefined => {
                        while next_positional < ordered.len() && ordered[next_positional].is_some()
                        {
                            next_positional += 1;
                        }
                        if next_positional < ordered.len() {
                            ordered[next_positional] = Some(value);
                            next_positional += 1;
                        } else {
                            extra.push(value);
                        }
                    }
                    _ => return Value::Undefined,
                }
            }
            let supplied = ordered.iter().rposition(Option::is_some);
            let mut bound = ordered
                .into_iter()
                .take(supplied.map_or(0, |index| index + 1))
                .map(|value| value.unwrap_or(Value::Undefined))
                .collect::<Vec<_>>();
            bound.extend(extra);
            ctx.invoke_with_receiver(constructor, Value::Null, &bound)
        }),
    );
}
