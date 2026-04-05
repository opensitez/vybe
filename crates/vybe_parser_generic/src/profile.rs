//! Language profile — compilation semantics per language.
//!
//! The grammar defines syntax (how to parse). The profile defines semantics
//! (how to compile). Together they fully describe a language.

use std::collections::HashMap;

/// Compilation semantics for a language.
#[derive(Debug, Clone)]
pub struct LanguageProfile {
    /// How functions return values.
    pub function_return: ReturnStyle,

    /// The keyword used for `self`/`this` in methods.
    pub self_keyword: String,

    /// Constructor method name (matched case-insensitively for case-insensitive languages).
    pub constructor_name: String,

    /// Whether method bodies are separate from class declarations.
    /// Pascal: true (type TFoo = class...end; then constructor TFoo.Create; begin..end;)
    /// Python/JS/Ruby: false (methods are inline in class body)
    pub separated_methods: bool,

    /// Whether bare field names in methods resolve to self.field.
    /// Pascal: true (FName → Self.FName)
    /// Python/JS: false (must write self.name / this.name explicitly)
    pub implicit_self_fields: bool,

    /// Whether the self parameter is explicit in method signatures.
    /// Python: true (def method(self, x))
    /// Pascal/JS/Ruby: false (self/this is implicit)
    pub explicit_self_param: bool,

    /// Whether enum values are compiled as global ordinal constants.
    pub enum_as_ordinals: bool,

    /// Builtin function mappings: source name → emission action.
    pub builtins: HashMap<String, BuiltinDef>,
}

/// How a function returns its value.
#[derive(Debug, Clone, PartialEq)]
pub enum ReturnStyle {
    /// Assign to a Result slot; function epilogue returns it. (Pascal, VB)
    ResultSlot,
    /// Explicit `return expr` statements. (Python, JS, C#, PHP)
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
    /// Print (variadic)
    Print,
    /// String length
    StrLength,
    /// Emit nothing (no-op, e.g. randomize, free)
    Noop,
}

impl LanguageProfile {
    pub fn pascal() -> Self {
        let mut builtins = HashMap::new();

        // I/O
        builtins.insert("writeln".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 0, max_args: 255 });
        builtins.insert("write".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 0, max_args: 255 });

        // Math (direct opcodes)
        builtins.insert("abs".into(), BuiltinDef { emit: BuiltinEmit::Opcode("abs".into()), min_args: 1, max_args: 1 });
        builtins.insert("sqrt".into(), BuiltinDef { emit: BuiltinEmit::Opcode("sqrt".into()), min_args: 1, max_args: 1 });
        builtins.insert("round".into(), BuiltinDef { emit: BuiltinEmit::Opcode("round".into()), min_args: 1, max_args: 1 });
        builtins.insert("trunc".into(), BuiltinDef { emit: BuiltinEmit::Opcode("trunc".into()), min_args: 1, max_args: 1 });
        builtins.insert("floor".into(), BuiltinDef { emit: BuiltinEmit::Opcode("floor".into()), min_args: 1, max_args: 1 });
        builtins.insert("ceil".into(), BuiltinDef { emit: BuiltinEmit::Opcode("ceil".into()), min_args: 1, max_args: 1 });
        builtins.insert("min".into(), BuiltinDef { emit: BuiltinEmit::Opcode("min".into()), min_args: 2, max_args: 2 });
        builtins.insert("max".into(), BuiltinDef { emit: BuiltinEmit::Opcode("max".into()), min_args: 2, max_args: 2 });
        builtins.insert("sqr".into(), BuiltinDef { emit: BuiltinEmit::Opcode("sqr".into()), min_args: 1, max_args: 1 });

        // Math (host calls)
        builtins.insert("power".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:math".into(), "pow".into()), min_args: 2, max_args: 2 });
        builtins.insert("sin".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:math".into(), "sin".into()), min_args: 1, max_args: 1 });
        builtins.insert("cos".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:math".into(), "cos".into()), min_args: 1, max_args: 1 });
        builtins.insert("exp".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:math".into(), "exp".into()), min_args: 1, max_args: 1 });
        builtins.insert("ln".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:math".into(), "log".into()), min_args: 1, max_args: 1 });
        builtins.insert("random".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:math".into(), "random".into()), min_args: 0, max_args: 0 });
        builtins.insert("randomize".into(), BuiltinDef { emit: BuiltinEmit::Noop, min_args: 0, max_args: 0 });

        // Succ/Pred
        builtins.insert("succ".into(), BuiltinDef { emit: BuiltinEmit::Opcode("succ".into()), min_args: 1, max_args: 1 });
        builtins.insert("pred".into(), BuiltinDef { emit: BuiltinEmit::Opcode("pred".into()), min_args: 1, max_args: 1 });

        // Inc/Dec (mutate variable)
        builtins.insert("inc".into(), BuiltinDef { emit: BuiltinEmit::MutateVar("add".into()), min_args: 1, max_args: 2 });
        builtins.insert("dec".into(), BuiltinDef { emit: BuiltinEmit::MutateVar("sub".into()), min_args: 1, max_args: 2 });

        // Strings
        builtins.insert("length".into(), BuiltinDef { emit: BuiltinEmit::StrLength, min_args: 1, max_args: 1 });
        builtins.insert("uppercase".into(), BuiltinDef { emit: BuiltinEmit::Opcode("to_upper".into()), min_args: 1, max_args: 1 });
        builtins.insert("upcase".into(), BuiltinDef { emit: BuiltinEmit::Opcode("to_upper".into()), min_args: 1, max_args: 1 });
        builtins.insert("lowercase".into(), BuiltinDef { emit: BuiltinEmit::Opcode("to_lower".into()), min_args: 1, max_args: 1 });
        builtins.insert("trim".into(), BuiltinDef { emit: BuiltinEmit::Opcode("trim".into()), min_args: 1, max_args: 1 });
        builtins.insert("concat".into(), BuiltinDef { emit: BuiltinEmit::Opcode("concat".into()), min_args: 1, max_args: 255 });
        builtins.insert("stringreplace".into(), BuiltinDef { emit: BuiltinEmit::Opcode("replace".into()), min_args: 3, max_args: 3 });
        builtins.insert("stringofchar".into(), BuiltinDef { emit: BuiltinEmit::Opcode("repeat".into()), min_args: 2, max_args: 2 });
        builtins.insert("leftstr".into(), BuiltinDef { emit: BuiltinEmit::Opcode("leftstr".into()), min_args: 2, max_args: 2 });

        // Type conversions
        builtins.insert("inttostr".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "toString".into()), min_args: 1, max_args: 1 });
        builtins.insert("floattostr".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "toString".into()), min_args: 1, max_args: 1 });
        builtins.insert("strtoint".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "parseInt".into()), min_args: 1, max_args: 1 });
        builtins.insert("strtointdef".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "parseInt".into()), min_args: 1, max_args: 1 });
        builtins.insert("strtofloat".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "parseFloat".into()), min_args: 1, max_args: 1 });
        builtins.insert("chr".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "chr".into()), min_args: 1, max_args: 1 });
        builtins.insert("ord".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "charCodeAt".into()), min_args: 1, max_args: 1 });

        // Array operations
        builtins.insert("high".into(), BuiltinDef { emit: BuiltinEmit::Opcode("high".into()), min_args: 1, max_args: 1 });
        builtins.insert("low".into(), BuiltinDef { emit: BuiltinEmit::Opcode("low".into()), min_args: 0, max_args: 1 });
        builtins.insert("setlength".into(), BuiltinDef { emit: BuiltinEmit::Opcode("setlength".into()), min_args: 2, max_args: 2 });
        builtins.insert("assigned".into(), BuiltinDef { emit: BuiltinEmit::Opcode("assigned".into()), min_args: 1, max_args: 1 });

        // Object lifecycle
        builtins.insert("freeandnil".into(), BuiltinDef { emit: BuiltinEmit::Opcode("freeandnil".into()), min_args: 1, max_args: 1 });
        builtins.insert("free".into(), BuiltinDef { emit: BuiltinEmit::Noop, min_args: 0, max_args: 1 });

        Self {
            function_return: ReturnStyle::ResultSlot,
            self_keyword: "Self".into(),
            constructor_name: "Create".into(),
            separated_methods: true,
            implicit_self_fields: true,
            explicit_self_param: false,
            enum_as_ordinals: true,
            builtins,
        }
    }

    pub fn python() -> Self {
        let mut builtins = HashMap::new();

        builtins.insert("print".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 0, max_args: 255 });
        builtins.insert("len".into(), BuiltinDef { emit: BuiltinEmit::StrLength, min_args: 1, max_args: 1 });
        builtins.insert("abs".into(), BuiltinDef { emit: BuiltinEmit::Opcode("abs".into()), min_args: 1, max_args: 1 });
        builtins.insert("round".into(), BuiltinDef { emit: BuiltinEmit::Opcode("round".into()), min_args: 1, max_args: 1 });
        builtins.insert("min".into(), BuiltinDef { emit: BuiltinEmit::Opcode("min".into()), min_args: 2, max_args: 2 });
        builtins.insert("max".into(), BuiltinDef { emit: BuiltinEmit::Opcode("max".into()), min_args: 2, max_args: 2 });
        builtins.insert("str".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "toString".into()), min_args: 1, max_args: 1 });
        builtins.insert("int".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "parseInt".into()), min_args: 1, max_args: 1 });
        builtins.insert("float".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "parseFloat".into()), min_args: 1, max_args: 1 });
        builtins.insert("range".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:collections".into(), "range".into()), min_args: 1, max_args: 3 });
        builtins.insert("isinstance".into(), BuiltinDef { emit: BuiltinEmit::Opcode("isinstance".into()), min_args: 2, max_args: 2 });
        builtins.insert("type".into(), BuiltinDef { emit: BuiltinEmit::Opcode("typeof".into()), min_args: 1, max_args: 1 });
        builtins.insert("sorted".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:array".into(), "sort".into()), min_args: 1, max_args: 1 });
        builtins.insert("reversed".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:array".into(), "reverse".into()), min_args: 1, max_args: 1 });
        builtins.insert("enumerate".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:collections".into(), "enumerate".into()), min_args: 1, max_args: 1 });
        builtins.insert("zip".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:collections".into(), "zip".into()), min_args: 2, max_args: 2 });
        builtins.insert("map".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:collections".into(), "map".into()), min_args: 2, max_args: 2 });
        builtins.insert("filter".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:collections".into(), "filter".into()), min_args: 2, max_args: 2 });
        builtins.insert("input".into(), BuiltinDef { emit: BuiltinEmit::HostCall("wasi:cli".into(), "readLine".into()), min_args: 0, max_args: 1 });

        Self {
            function_return: ReturnStyle::Explicit,
            self_keyword: "self".into(),
            constructor_name: "__init__".into(),
            separated_methods: false,
            implicit_self_fields: false,
            explicit_self_param: true,
            enum_as_ordinals: false,
            builtins,
        }
    }

    pub fn vb() -> Self {
        let mut builtins = HashMap::new();

        builtins.insert("console.writeline".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 0, max_args: 255 });
        builtins.insert("msgbox".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 1, max_args: 1 });
        builtins.insert("len".into(), BuiltinDef { emit: BuiltinEmit::StrLength, min_args: 1, max_args: 1 });
        builtins.insert("cstr".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "toString".into()), min_args: 1, max_args: 1 });
        builtins.insert("cint".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "parseInt".into()), min_args: 1, max_args: 1 });
        builtins.insert("cdbl".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "parseFloat".into()), min_args: 1, max_args: 1 });
        builtins.insert("ucase".into(), BuiltinDef { emit: BuiltinEmit::Opcode("to_upper".into()), min_args: 1, max_args: 1 });
        builtins.insert("lcase".into(), BuiltinDef { emit: BuiltinEmit::Opcode("to_lower".into()), min_args: 1, max_args: 1 });
        builtins.insert("trim".into(), BuiltinDef { emit: BuiltinEmit::Opcode("trim".into()), min_args: 1, max_args: 1 });
        builtins.insert("abs".into(), BuiltinDef { emit: BuiltinEmit::Opcode("abs".into()), min_args: 1, max_args: 1 });
        builtins.insert("math.sqrt".into(), BuiltinDef { emit: BuiltinEmit::Opcode("sqrt".into()), min_args: 1, max_args: 1 });
        builtins.insert("math.round".into(), BuiltinDef { emit: BuiltinEmit::Opcode("round".into()), min_args: 1, max_args: 1 });
        builtins.insert("math.floor".into(), BuiltinDef { emit: BuiltinEmit::Opcode("floor".into()), min_args: 1, max_args: 1 });
        builtins.insert("math.max".into(), BuiltinDef { emit: BuiltinEmit::Opcode("max".into()), min_args: 2, max_args: 2 });
        builtins.insert("math.min".into(), BuiltinDef { emit: BuiltinEmit::Opcode("min".into()), min_args: 2, max_args: 2 });

        // String functions
        builtins.insert("left".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "left".into()), min_args: 2, max_args: 2 });
        builtins.insert("right".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "right".into()), min_args: 2, max_args: 2 });
        builtins.insert("mid".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "mid".into()), min_args: 2, max_args: 3 });
        builtins.insert("instr".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "instr".into()), min_args: 2, max_args: 3 });
        builtins.insert("replace".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "replace".into()), min_args: 3, max_args: 3 });
        builtins.insert("split".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "split".into()), min_args: 1, max_args: 2 });
        builtins.insert("join".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "join".into()), min_args: 1, max_args: 2 });
        builtins.insert("ltrim".into(), BuiltinDef { emit: BuiltinEmit::Opcode("trim_start".into()), min_args: 1, max_args: 1 });
        builtins.insert("rtrim".into(), BuiltinDef { emit: BuiltinEmit::Opcode("trim_end".into()), min_args: 1, max_args: 1 });
        builtins.insert("space".into(), BuiltinDef { emit: BuiltinEmit::Opcode("space".into()), min_args: 1, max_args: 1 });
        builtins.insert("asc".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "charCodeAt".into()), min_args: 1, max_args: 1 });
        builtins.insert("chr".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:string".into(), "fromCharCode".into()), min_args: 1, max_args: 1 });
        // Type checking
        builtins.insert("isnumeric".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "isNumeric".into()), min_args: 1, max_args: 1 });
        builtins.insert("isnothing".into(), BuiltinDef { emit: BuiltinEmit::Opcode("is_null".into()), min_args: 1, max_args: 1 });
        builtins.insert("val".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "parseFloat".into()), min_args: 1, max_args: 1 });
        // Array
        builtins.insert("ubound".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:array".into(), "ubound".into()), min_args: 1, max_args: 1 });
        // Math
        builtins.insert("math.abs".into(), BuiltinDef { emit: BuiltinEmit::Opcode("abs".into()), min_args: 1, max_args: 1 });
        builtins.insert("math.ceiling".into(), BuiltinDef { emit: BuiltinEmit::Opcode("ceil".into()), min_args: 1, max_args: 1 });
        builtins.insert("math.pow".into(), BuiltinDef { emit: BuiltinEmit::Opcode("pow".into()), min_args: 2, max_args: 2 });
        builtins.insert("math.log".into(), BuiltinDef { emit: BuiltinEmit::Opcode("log".into()), min_args: 1, max_args: 1 });
        builtins.insert("math.sin".into(), BuiltinDef { emit: BuiltinEmit::Opcode("sin".into()), min_args: 1, max_args: 1 });
        builtins.insert("math.cos".into(), BuiltinDef { emit: BuiltinEmit::Opcode("cos".into()), min_args: 1, max_args: 1 });

        Self {
            function_return: ReturnStyle::ResultSlot,
            self_keyword: "Me".into(),
            constructor_name: "New".into(),
            separated_methods: false,
            implicit_self_fields: true,
            explicit_self_param: false,
            enum_as_ordinals: true,
            builtins,
        }
    }

    pub fn javascript() -> Self {
        let mut builtins = HashMap::new();

        builtins.insert("console.log".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 0, max_args: 255 });
        builtins.insert("alert".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 1, max_args: 1 });
        builtins.insert("parseint".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "parseInt".into()), min_args: 1, max_args: 1 });
        builtins.insert("parsefloat".into(), BuiltinDef { emit: BuiltinEmit::HostCall("vybe:convert".into(), "parseFloat".into()), min_args: 1, max_args: 1 });

        Self {
            function_return: ReturnStyle::Explicit,
            self_keyword: "this".into(),
            constructor_name: "constructor".into(),
            separated_methods: false,
            implicit_self_fields: false,
            explicit_self_param: false,
            enum_as_ordinals: false,
            builtins,
        }
    }

    /// Look up a builtin by name (case-insensitive for case-insensitive languages).
    pub fn lookup_builtin(&self, name: &str, case_sensitive: bool) -> Option<&BuiltinDef> {
        if case_sensitive {
            self.builtins.get(name)
        } else {
            let lower = name.to_lowercase();
            self.builtins.get(&lower)
        }
    }
}
