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
    pub value_methods: HashMap<String, BuiltinDef>,

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

    /// Look up a value method (instance method on a value: str.toUpperCase(), arr.push()).
    pub fn lookup_value_method(&self, name: &str) -> Option<&BuiltinDef> {
        if self.case_sensitive {
            self.value_methods.get(name)
        } else {
            self.value_methods.get(&name.to_lowercase())
        }
    }

    /// Look up a module alias (JS: console → wasi:cli, Math → vybe:math).
    pub fn lookup_module_alias(&self, name: &str) -> Option<&str> {
        self.module_aliases.get(name).map(|s| s.as_str())
    }

    /// Look up a namespace constant value (Math.PI, Number.MAX_SAFE_INTEGER).
    pub fn lookup_constant(&self, name: &str) -> Option<&ConstantValue> {
        self.namespace_constants.get(name)
    }

    /// Look up an array method routing (map → __array_map).
    pub fn lookup_array_method(&self, name: &str) -> Option<&str> {
        self.array_methods.get(name).map(|s| s.as_str())
    }
}
