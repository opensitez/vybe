//! vybex — Universal compiler for VB and JS (more languages to come).
//!
//! Usage: vybex <file.vb|file.js>
//!
//! Parses source with pest grammar, compiles via common AST + profile, runs on VM.

use std::path::Path;
use vybe_bytecode::VM;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: vybex <file.vb|file.js|file.pas>");
        std::process::exit(1);
    }

    let path = Path::new(&args[1]);
    let source = match std::fs::read_to_string(path) {
        Ok(s) => s,
        Err(e) => { eprintln!("Cannot read {}: {}", path.display(), e); std::process::exit(1); }
    };

    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");
    let lang = match ext {
        "vb" => "vb",
        "js" | "mjs" => "js",
        "pas" | "pp" | "dpr" | "lpr" => "pascal",
        _ => { eprintln!("Unknown file extension: .{}", ext); std::process::exit(1); }
    };

    // Parse source → common AST
    let module = match lang {
        "vb" => match vybex::languages::vb::parse(&source) {
            Ok(m) => m,
            Err(e) => { eprintln!("Parse error: {}", e); std::process::exit(1); }
        },
        "js" => match vybex::languages::js::parse(&source) {
            Ok(m) => m,
            Err(e) => { eprintln!("Parse error: {}", e); std::process::exit(1); }
        },
        "pascal" => match vybex::languages::pascal::parse(&source) {
            Ok(m) => m,
            Err(e) => { eprintln!("Parse error: {}", e); std::process::exit(1); }
        },
        _ => unreachable!(),
    };

    // Load profile
    let profile = match load_profile(lang) {
        Ok(p) => p,
        Err(e) => { eprintln!("Profile error: {}", e); std::process::exit(1); }
    };

    // Compile AST → bytecode
    let chunks = match vybex::compiler::Compiler::with_profile(profile).compile(&module) {
        Ok(c) => c,
        Err(e) => { eprintln!("Compile error: {}", e); std::process::exit(1); }
    };

    // Run on VM
    let mut vm = VM::new();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {}", e); std::process::exit(1); }
    }

    if gui.lock().unwrap().should_run {
        vybe_cli::runner::launch_vm_form(vm, gui, None);
    }
}

/// Load a language profile from the profile file next to the grammar.
fn load_profile(lang: &str) -> Result<vybex::profile::LanguageProfile, String> {
    // Look for profile relative to the binary's working directory
    let candidates = [
        format!("crates/vybex/src/languages/{}/profile", lang),
        format!("crates/newparser/src/languages/{}/profile", lang),
        format!("src/languages/{}/profile", lang),
        format!("languages/{}/profile", lang),
    ];
    for path in &candidates {
        if Path::new(path).exists() {
            let src = std::fs::read_to_string(path)
                .map_err(|e| format!("Cannot read profile {}: {}", path, e))?;
            return parse_profile(&src);
        }
    }
    Err(format!("Profile not found for '{}'. Looked in {:?}", lang, candidates))
}

/// Parse a TOML profile into a LanguageProfile.
fn parse_profile(src: &str) -> Result<vybex::profile::LanguageProfile, String> {
    use vybex::profile::*;
    use std::collections::HashMap;
    use toml::Value;

    let root: Value = toml::from_str(src)
        .map_err(|e| format!("TOML parse error in profile: {}", e))?;

    let compiler = root.get("compiler")
        .ok_or("Missing [compiler] section")?;

    let function_return = match compiler.get("function_return")
        .and_then(|v| v.as_str()).unwrap_or("explicit") {
        "result_slot" => ReturnStyle::ResultSlot,
        "last_expression" => ReturnStyle::LastExpression,
        _ => ReturnStyle::Explicit,
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
        _ => StringIndexing::ZeroBased,
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

    fn parse_builtin_table(root: &Value, section: &str) -> HashMap<String, BuiltinDef> {
        let mut map = HashMap::new();
        if let Some(bt) = root.get(section).and_then(|v| v.as_table()) {
            for (name, val) in bt {
                if let Some(t) = val.as_table() {
                    let emit_str = t.get("emit").and_then(|v| v.as_str()).unwrap_or("noop");
                    let min_args = t.get("min_args").and_then(|v| v.as_integer()).unwrap_or(0) as u8;
                    let max_args = t.get("max_args").and_then(|v| v.as_integer()).unwrap_or(255) as u8;
                    if let Some(emit) = parse_emit(emit_str) {
                        map.insert(name.clone(), BuiltinDef { emit, min_args, max_args });
                    }
                }
            }
        }
        map
    }

    fn parse_emit(s: &str) -> Option<BuiltinEmit> {
        match s {
            "print" => Some(BuiltinEmit::Print),
            "str_length" => Some(BuiltinEmit::StrLength),
            "noop" => Some(BuiltinEmit::Noop),
            _ if s.starts_with("host:") => {
                let parts: Vec<&str> = s["host:".len()..].splitn(3, ':').collect();
                if parts.len() == 3 {
                    Some(BuiltinEmit::HostCall(format!("{}:{}", parts[0], parts[1]), parts[2].to_string()))
                } else { None }
            }
            _ if s.starts_with("opcode:") => Some(BuiltinEmit::Opcode(s["opcode:".len()..].to_string())),
            _ if s.starts_with("mutate:") => Some(BuiltinEmit::MutateVar(s["mutate:".len()..].to_string())),
            _ if s.starts_with("intrinsic:") => Some(BuiltinEmit::Intrinsic(s["intrinsic:".len()..].to_string())),
            _ => None,
        }
    }

    fn parse_string_table(root: &Value, section: &str) -> HashMap<String, String> {
        let mut map = HashMap::new();
        if let Some(t) = root.get(section).and_then(|v| v.as_table()) {
            for (k, v) in t {
                if let Some(s) = v.as_str() { map.insert(k.clone(), s.to_string()); }
            }
        }
        map
    }

    let builtins = parse_builtin_table(&root, "builtins");
    let value_methods = parse_builtin_table(&root, "value_methods");
    let intrinsics = parse_string_table(&root, "intrinsics");
    let module_aliases = parse_string_table(&root, "module_aliases");
    let array_methods = parse_string_table(&root, "array_methods");

    // Namespace config
    let namespaces = if let Some(ns) = root.get("namespaces") {
        NamespaceConfig {
            use_dotnet: ns.get("use_dotnet").and_then(|v| v.as_bool()).unwrap_or(false),
            extra_imports: ns.get("extra_imports").and_then(|v| v.as_array())
                .map(|a| a.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
                .unwrap_or_default(),
            roots: ns.get("roots").and_then(|v| v.as_array())
                .map(|a| a.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
                .unwrap_or_default(),
            default_imports: ns.get("default_imports").and_then(|v| v.as_array())
                .map(|a| a.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
                .unwrap_or_default(),
            constants: ns.get("constants").and_then(|v| v.as_array())
                .map(|a| a.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
                .unwrap_or_default(),
        }
    } else {
        NamespaceConfig::default()
    };

    // Known types
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

    // Namespace constants
    let mut namespace_constants = HashMap::new();
    if let Some(nc) = root.get("namespace_constants").and_then(|v| v.as_table()) {
        for (name, val) in nc {
            match val {
                Value::Float(f) => { namespace_constants.insert(name.clone(), ConstantValue::Float(*f)); }
                Value::Integer(i) => { namespace_constants.insert(name.clone(), ConstantValue::Float(*i as f64)); }
                Value::String(s) => {
                    if let Ok(f) = s.parse::<f64>() {
                        namespace_constants.insert(name.clone(), ConstantValue::Float(f));
                    } else {
                        namespace_constants.insert(name.clone(), ConstantValue::Str(s.clone()));
                    }
                }
                _ => {}
            }
        }
    }

    Ok(LanguageProfile {
        function_return, self_keyword, base_keyword, constructor_name,
        separated_methods, implicit_self_fields, explicit_self_param,
        enum_as_ordinals, case_sensitive, string_indexing,
        array_upper_bound_inclusive, parens_for_index, entry_point,
        hoist_var, dynamic_add, commonjs_require,
        partial_classes, byref_boxing, with_block,
        event_side_effects, new_with_initializer, new_from_initializer, linq_queries,
        builtins, intrinsics, namespaces, known_types,
        value_methods, module_aliases, namespace_constants, array_methods,
    })
}
