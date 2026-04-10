//! Language profile — compilation semantics per language.
//!
//! The grammar defines syntax (how to parse). The profile defines semantics
//! (how to compile). Together they fully describe a language.
//!
//! Profiles are loaded from `languages/<lang>/profile` files — no hardcoded
//! language knowledge lives in Rust code.

use std::collections::HashMap;

/// Compilation semantics for a language.
#[derive(Debug, Clone)]
pub struct LanguageProfile {
    /// How functions return values.
    pub function_return: ReturnStyle,

    /// The keyword used for `self`/`this` in methods.
    pub self_keyword: String,

    /// The keyword for base/super class reference (e.g. "mybase", "super", "inherited").
    pub base_keyword: Option<String>,

    /// Constructor method name (matched case-insensitively for case-insensitive languages).
    pub constructor_name: String,

    /// Whether method bodies are separate from class declarations.
    pub separated_methods: bool,

    /// Whether bare field names in methods resolve to self.field.
    pub implicit_self_fields: bool,

    /// Whether the self parameter is explicit in method signatures.
    pub explicit_self_param: bool,

    /// Whether enum values are compiled as global ordinal constants.
    pub enum_as_ordinals: bool,

    /// Whether the language is case-sensitive.
    pub case_sensitive: bool,

    /// String indexing: "zero_based" or "one_based" (VB).
    pub string_indexing: StringIndexing,

    /// Whether array upper bounds are inclusive (VB: Dim arr(5) = 6 elements).
    pub array_upper_bound_inclusive: bool,

    /// Whether parens are used for both calls and indexing (VB: arr(i)).
    pub parens_for_index: bool,

    /// Entry point function name to auto-call if defined (e.g. "main").
    pub entry_point: Option<String>,

    /// JS: `var` declarations are hoisted to function scope.
    pub hoist_var: bool,

    /// JS: `+` operator uses dynamic add (string concat if either operand is string).
    pub dynamic_add: bool,

    /// JS: support `require()` for CommonJS module loading.
    pub commonjs_require: bool,

    /// VB: merge Partial Class declarations before compiling.
    pub partial_classes: bool,

    /// VB: ByRef args wrapped in single-element arrays for call-by-reference.
    pub byref_boxing: bool,

    /// VB: With obj ... End With — bare .Member resolves to With target.
    pub with_block: bool,

    /// VB: member assignment on controls triggers property change side effects.
    pub event_side_effects: bool,

    /// VB: New Foo() With { .Prop = val } initializer syntax.
    pub new_with_initializer: bool,

    /// VB: New List(Of T) From { items } initializer syntax.
    pub new_from_initializer: bool,

    /// VB: LINQ query syntax compiled to method chains.
    pub linq_queries: bool,

    /// Builtin function mappings: source name → emission action.
    pub builtins: HashMap<String, BuiltinDef>,

    /// Multi-opcode intrinsic definitions referenced by `emit = "intrinsic:<name>"`.
    pub intrinsics: HashMap<String, String>,

    /// Namespace resolution config.
    pub namespaces: NamespaceConfig,

    /// Known types: name → (host_module, host_function) for New TypeName().
    pub known_types: HashMap<String, (String, String)>,

    /// Value methods: instance methods called on values (str.toUpperCase(), arr.push()).
    /// The object is passed as first arg to the host function.
    /// Value methods can have multiple overloads by arity.
    /// E.g. `Add(item)` for list (1 arg) vs `Add(key, value)` for dict (2 args).
    pub value_methods: HashMap<String, Vec<BuiltinDef>>,

    /// Module aliases: JS namespace objects → host modules (console → wasi:cli, Math → vybe:math).
    pub module_aliases: HashMap<String, String>,

    /// Namespace constants: property access that returns a value, NOT a function call.
    /// "Math.PI" → 3.14159..., "Number.MAX_SAFE_INTEGER" → 9007199254740991
    pub namespace_constants: HashMap<String, ConstantValue>,

    /// Array higher-order methods routed to compiled JS builtins.
    /// "map" → "__array_map", "filter" → "__array_filter", etc.
    pub array_methods: HashMap<String, String>,
}

/// A compile-time constant value.
#[derive(Debug, Clone)]
pub enum ConstantValue {
    Float(f64),
    Str(String),
}

/// String indexing style.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum StringIndexing {
    ZeroBased,
    OneBased,
}

/// Namespace resolution configuration.
#[derive(Debug, Clone, Default)]
pub struct NamespaceConfig {
    /// Use .NET BCL resolution from vybe_compiler_common::dotnet.
    /// When true, the compiler uses dotnet::namespace_roots(), dotnet::default_interface_imports(),
    /// dotnet::resolve_dotted_name(), etc. — the full .NET resolution pipeline.
    pub use_dotnet: bool,
    /// Additional imports beyond the defaults (e.g. "microsoft.visualbasic" for VB).
    pub extra_imports: Vec<String>,
    /// Known namespace roots (used when use_dotnet is false).
    pub roots: Vec<String>,
    /// Default imports always available (used when use_dotnet is false).
    pub default_imports: Vec<String>,
    /// Known constants (property access, not function call).
    pub constants: Vec<String>,
}

/// How a function returns its value.
#[derive(Debug, Clone, PartialEq)]
pub enum ReturnStyle {
    /// Assign to a Result slot; function epilogue returns it. (Pascal)
    ResultSlot,
    /// Explicit `return expr` statements. (Python, JS, C#, PHP, VB)
    Explicit,
    /// Last expression in body is the return value. (Ruby)
    LastExpression,
}

/// Definition of a builtin function's compilation.
#[derive(Debug, Clone)]
pub struct BuiltinDef {
    pub emit: BuiltinEmit,
    pub min_args: u8,
    pub max_args: u8,
}

/// What to emit for a builtin call.
#[derive(Debug, Clone)]
pub enum BuiltinEmit {
    /// Call a host import: (module, name)
    HostCall(String, String),
    /// Emit a direct opcode sequence by name.
    Opcode(String),
    /// Mutate a variable: var = var OP arg. (Inc, Dec)
    MutateVar(String),  // "add" or "sub"
    /// Multi-opcode intrinsic: name references [intrinsics] table in profile.
    Intrinsic(String),
    /// Dispatch to a compiler_common opcode-style emitter (args already on stack).
    /// e.g. "dict.set_dynamic", "collections.push", "strings.length"
    Common(String),
    /// Call a stdlib function (e.g. __vybe_sorted) via global_get + call_ref.
    /// Func ref is pushed BEFORE args.
    /// Name is the operation, e.g. "sorted", "range", "sum", "min", "max"
    Stdlib(String),
    /// Print (variadic)
    Print,
    /// String length
    StrLength,
    /// Emit nothing (no-op, e.g. randomize, free)
    Noop,
}

impl LanguageProfile {
    /// Look up a builtin by name (case-insensitive for case-insensitive languages).
    pub fn lookup_builtin(&self, name: &str) -> Option<&BuiltinDef> {
        if self.case_sensitive {
            self.builtins.get(name)
        } else {
            let lower = name.to_lowercase();
            self.builtins.get(&lower)
        }
    }

    /// Look up a known type constructor mapping.
    pub fn lookup_known_type(&self, name: &str) -> Option<(&str, &str)> {
        let key = if self.case_sensitive { name.to_string() } else { name.to_lowercase() };
        self.known_types.get(&key).map(|(m, f)| (m.as_str(), f.as_str()))
    }

    /// Check if a name is a known namespace root.
    pub fn is_namespace_root(&self, name: &str) -> bool {
        let key = if self.case_sensitive { name.to_string() } else { name.to_lowercase() };
        self.namespaces.roots.iter().any(|r| r == &key)
    }

    /// Check if a name is a known constant (property access, not call).
    pub fn is_namespace_constant(&self, name: &str) -> bool {
        let key = if self.case_sensitive { name.to_string() } else { name.to_lowercase() };
        self.namespaces.constants.iter().any(|c| c == &key)
    }

    /// Look up a value method by name + arity.
    /// Returns the first overload whose arity range matches.
    pub fn lookup_value_method(&self, name: &str, argc: u8) -> Option<&BuiltinDef> {
        let key = if self.case_sensitive { name.to_string() } else { name.to_lowercase() };
        let overloads = self.value_methods.get(&key)?;
        overloads.iter().find(|d| argc >= d.min_args && argc <= d.max_args)
            .or_else(|| overloads.first()) // fallback: first overload if no arity match
    }

    /// Check if a value method exists by name (any arity).
    pub fn has_value_method(&self, name: &str) -> bool {
        let key = if self.case_sensitive { name.to_string() } else { name.to_lowercase() };
        self.value_methods.contains_key(&key)
    }

    /// Look up a module alias (JS: console → wasi:cli, Math → vybe:math).
    pub fn lookup_module_alias(&self, name: &str) -> Option<&str> {
        self.module_aliases.get(name).map(|s| s.as_str())
    }

    /// Look up a namespace constant value (Math.PI, Number.MAX_SAFE_INTEGER).
    pub fn lookup_constant(&self, name: &str) -> Option<&ConstantValue> {
        if self.case_sensitive {
            self.namespace_constants.get(name)
        } else {
            self.namespace_constants.get(&name.to_lowercase())
        }
    }

    /// Look up an array method routing (map → __array_map).
    pub fn lookup_array_method(&self, name: &str) -> Option<&str> {
        self.array_methods.get(name).map(|s| s.as_str())
    }
}

/// Parse a TOML profile source into a LanguageProfile.
pub fn parse_profile(src: &str) -> Result<LanguageProfile, String> {
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
            _ if s.starts_with("common:") => Some(BuiltinEmit::Common(s["common:".len()..].to_string())),
            _ if s.starts_with("stdlib:") => Some(BuiltinEmit::Stdlib(s["stdlib:".len()..].to_string())),
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

    fn parse_value_methods_table(root: &Value) -> HashMap<String, Vec<BuiltinDef>> {
        let mut map: HashMap<String, Vec<BuiltinDef>> = HashMap::new();
        if let Some(bt) = root.get("value_methods").and_then(|v| v.as_table()) {
            for (name, val) in bt {
                // Either a single inline table or an array of inline tables (overloads)
                if let Some(arr) = val.as_array() {
                    for entry in arr {
                        if let Some(t) = entry.as_table() {
                            let emit_str = t.get("emit").and_then(|v| v.as_str()).unwrap_or("noop");
                            let min_args = t.get("min_args").and_then(|v| v.as_integer()).unwrap_or(0) as u8;
                            let max_args = t.get("max_args").and_then(|v| v.as_integer()).unwrap_or(255) as u8;
                            if let Some(emit) = parse_emit(emit_str) {
                                map.entry(name.clone()).or_default().push(BuiltinDef { emit, min_args, max_args });
                            }
                        }
                    }
                } else if let Some(t) = val.as_table() {
                    let emit_str = t.get("emit").and_then(|v| v.as_str()).unwrap_or("noop");
                    let min_args = t.get("min_args").and_then(|v| v.as_integer()).unwrap_or(0) as u8;
                    let max_args = t.get("max_args").and_then(|v| v.as_integer()).unwrap_or(255) as u8;
                    if let Some(emit) = parse_emit(emit_str) {
                        map.entry(name.clone()).or_default().push(BuiltinDef { emit, min_args, max_args });
                    }
                }
            }
        }
        map
    }

    let builtins = parse_builtin_table(&root, "builtins");
    let value_methods = parse_value_methods_table(&root);
    let intrinsics = parse_string_table(&root, "intrinsics");
    let module_aliases = parse_string_table(&root, "module_aliases");
    let array_methods = parse_string_table(&root, "array_methods");

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
