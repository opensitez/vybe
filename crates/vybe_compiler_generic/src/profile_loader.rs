//! Load a LanguageProfile from a `languages/<lang>/profile` file.
//!
//! The file format is TOML. See `languages/vb/profile` for the canonical example.

use std::collections::HashMap;
use toml::Value;
use vybe_parser_generic::profile::{self, *};

/// Load a profile by language name.
/// Looks for `languages/<lang>/profile`.
pub fn load_profile_for(lang: &str) -> Result<LanguageProfile, String> {
    let candidates = [
        format!("languages/{}/profile", lang),
        format!("../../languages/{}/profile", lang),
    ];
    for path in &candidates {
        if std::path::Path::new(path).exists() {
            let src = std::fs::read_to_string(path)
                .map_err(|e| format!("Cannot read profile {}: {}", path, e))?;
            return parse_profile(&src);
        }
    }
    Err(format!("Profile file not found for language '{}'. Looked in languages/{}/profile", lang, lang))
}

pub fn parse_profile(src: &str) -> Result<LanguageProfile, String> {
    let root: Value = toml::from_str(src)
        .map_err(|e| format!("TOML parse error in profile: {}", e))?;

    let compiler = root.get("compiler")
        .ok_or("Missing [compiler] section")?;

    let function_return = match compiler.get("function_return")
        .and_then(|v| v.as_str()).unwrap_or("explicit") {
        "result_slot"      => ReturnStyle::ResultSlot,
        "last_expression"  => ReturnStyle::LastExpression,
        _                  => ReturnStyle::Explicit,
    };

    let self_keyword = compiler.get("self_keyword")
        .and_then(|v| v.as_str()).unwrap_or("this").to_string();
    let base_keyword = compiler.get("base_keyword")
        .and_then(|v| v.as_str()).map(|s| s.to_string());
    let constructor_name = compiler.get("constructor_name")
        .and_then(|v| v.as_str()).unwrap_or("constructor").to_string();
    let separated_methods = compiler.get("separated_methods")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let implicit_self_fields = compiler.get("implicit_self_fields")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let explicit_self_param = compiler.get("explicit_self_param")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let enum_as_ordinals = compiler.get("enum_as_ordinals")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let case_sensitive = compiler.get("case_sensitive")
        .and_then(|v| v.as_bool()).unwrap_or(true);
    let string_indexing = match compiler.get("string_indexing")
        .and_then(|v| v.as_str()).unwrap_or("zero_based") {
        "one_based" => StringIndexing::OneBased,
        _           => StringIndexing::ZeroBased,
    };
    let array_upper_bound_inclusive = compiler.get("array_upper_bound_inclusive")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let parens_for_index = compiler.get("parens_for_index")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let entry_point = compiler.get("entry_point")
        .and_then(|v| v.as_str()).map(|s| s.to_string());
    let hoist_var = compiler.get("hoist_var")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let dynamic_add = compiler.get("dynamic_add")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let commonjs_require = compiler.get("commonjs_require")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let partial_classes = compiler.get("partial_classes")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let byref_boxing = compiler.get("byref_boxing")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let with_block = compiler.get("with_block")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let event_side_effects = compiler.get("event_side_effects")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let new_with_initializer = compiler.get("new_with_initializer")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let new_from_initializer = compiler.get("new_from_initializer")
        .and_then(|v| v.as_bool()).unwrap_or(false);
    let linq_queries = compiler.get("linq_queries")
        .and_then(|v| v.as_bool()).unwrap_or(false);

    // Parse builtins
    let mut builtins = HashMap::new();
    if let Some(bt) = root.get("builtins").and_then(|v| v.as_table()) {
        for (name, val) in bt {
            if let Some(def) = parse_builtin_def(val) {
                builtins.insert(name.clone(), def);
            }
        }
    }

    // Parse intrinsics
    let mut intrinsics = HashMap::new();
    if let Some(it) = root.get("intrinsics").and_then(|v| v.as_table()) {
        for (name, val) in it {
            if let Some(s) = val.as_str() {
                intrinsics.insert(name.clone(), s.to_string());
            }
        }
    }

    // Parse namespaces
    let namespaces = parse_namespaces(&root);

    // Parse known types
    let mut known_types = HashMap::new();
    if let Some(kt) = root.get("known_types").and_then(|v| v.as_table()) {
        for (name, val) in kt {
            if let Some(arr) = val.as_array() {
                if arr.len() == 2 {
                    if let (Some(m), Some(f)) = (arr[0].as_str(), arr[1].as_str()) {
                        known_types.insert(name.clone(), (m.to_string(), f.to_string()));
                    }
                }
            }
        }
    }

    // Parse value methods
    let mut value_methods = HashMap::new();
    if let Some(vm) = root.get("value_methods").and_then(|v| v.as_table()) {
        for (name, val) in vm {
            if let Some(def) = parse_builtin_def(val) {
                value_methods.insert(name.clone(), def);
            }
        }
    }

    // Parse module aliases
    let mut module_aliases = HashMap::new();
    if let Some(ma) = root.get("module_aliases").and_then(|v| v.as_table()) {
        for (name, val) in ma {
            if let Some(s) = val.as_str() {
                module_aliases.insert(name.clone(), s.to_string());
            }
        }
    }

    // Parse namespace constants
    let mut namespace_constants = HashMap::new();
    if let Some(nc) = root.get("namespace_constants").and_then(|v| v.as_table()) {
        for (name, val) in nc {
            let cv = match val {
                Value::Float(f) => Some(ConstantValue::Float(*f)),
                Value::Integer(i) => Some(ConstantValue::Float(*i as f64)),
                Value::String(s) => {
                    // Try parse as float, otherwise store as string
                    if let Ok(f) = s.parse::<f64>() {
                        Some(ConstantValue::Float(f))
                    } else {
                        Some(ConstantValue::Str(s.clone()))
                    }
                }
                _ => None,
            };
            if let Some(c) = cv {
                namespace_constants.insert(name.clone(), c);
            }
        }
    }

    // Parse array methods
    let mut array_methods = HashMap::new();
    if let Some(am) = root.get("array_methods").and_then(|v| v.as_table()) {
        for (name, val) in am {
            if let Some(s) = val.as_str() {
                array_methods.insert(name.clone(), s.to_string());
            }
        }
    }

    Ok(LanguageProfile {
        function_return,
        self_keyword,
        base_keyword,
        constructor_name,
        separated_methods,
        implicit_self_fields,
        explicit_self_param,
        enum_as_ordinals,
        case_sensitive,
        string_indexing,
        array_upper_bound_inclusive,
        parens_for_index,
        entry_point,
        hoist_var,
        dynamic_add,
        commonjs_require,
        partial_classes,
        byref_boxing,
        with_block,
        event_side_effects,
        new_with_initializer,
        new_from_initializer,
        linq_queries,
        builtins,
        intrinsics,
        namespaces,
        known_types,
        value_methods,
        module_aliases,
        namespace_constants,
        array_methods,
    })
}

fn parse_namespaces(root: &Value) -> NamespaceConfig {
    let ns = match root.get("namespaces") {
        Some(v) => v,
        None => return NamespaceConfig::default(),
    };

    let use_dotnet = ns.get("use_dotnet")
        .and_then(|v| v.as_bool()).unwrap_or(false);

    let extra_imports = ns.get("extra_imports")
        .and_then(|v| v.as_array())
        .map(|arr| arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
        .unwrap_or_default();

    let roots = ns.get("roots")
        .and_then(|v| v.as_array())
        .map(|arr| arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
        .unwrap_or_default();

    let default_imports = ns.get("default_imports")
        .and_then(|v| v.as_array())
        .map(|arr| arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
        .unwrap_or_default();

    let constants = ns.get("constants")
        .and_then(|v| v.as_array())
        .map(|arr| arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
        .unwrap_or_default();

    NamespaceConfig { use_dotnet, extra_imports, roots, default_imports, constants }
}

fn parse_builtin_def(val: &Value) -> Option<BuiltinDef> {
    let t = val.as_table()?;
    let emit_str = t.get("emit").and_then(|v| v.as_str()).unwrap_or("noop");
    let min_args = t.get("min_args").and_then(|v| v.as_integer()).unwrap_or(0) as u8;
    let max_args = t.get("max_args").and_then(|v| v.as_integer()).unwrap_or(255) as u8;

    let emit = parse_builtin_emit(emit_str)?;
    Some(BuiltinDef { emit, min_args, max_args })
}

fn parse_builtin_emit(s: &str) -> Option<BuiltinEmit> {
    match s {
        "print"      => Some(BuiltinEmit::Print),
        "str_length" => Some(BuiltinEmit::StrLength),
        "noop"       => Some(BuiltinEmit::Noop),
        _ if s.starts_with("host:") => {
            // "host:wasi:cli:log" → module="wasi:cli", func="log"
            let parts: Vec<&str> = s["host:".len()..].splitn(3, ':').collect();
            if parts.len() == 3 {
                Some(BuiltinEmit::HostCall(
                    format!("{}:{}", parts[0], parts[1]),
                    parts[2].to_string(),
                ))
            } else {
                None
            }
        }
        _ if s.starts_with("opcode:") => {
            Some(BuiltinEmit::Opcode(s["opcode:".len()..].to_string()))
        }
        _ if s.starts_with("mutate:") => {
            Some(BuiltinEmit::MutateVar(s["mutate:".len()..].to_string()))
        }
        _ if s.starts_with("intrinsic:") => {
            Some(BuiltinEmit::Intrinsic(s["intrinsic:".len()..].to_string()))
        }
        _ => None,
    }
}
