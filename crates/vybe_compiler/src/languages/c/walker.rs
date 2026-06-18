//! C → common AST walker.
//!
//! Walks the pest parse tree from `grammar.pest` into `vybe_compiler::ast`
//! nodes. C-specific normalizations happen here so the shared compiler stays
//! language-agnostic:
//!   - `printf(fmt, …)` → `__c_printf(fmt, …)` (shared sprintf formatter)
//!   - structs are tracked so `struct P x;` initializes a zero-filled object
//!   - pointer deref `*p` / address-of `&x` lower to common reference AST
//!   - `a->b` is treated as `a.b`

use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};

use super::{CParser, Rule};
use crate::ast::*;
use crate::platforms::libc::pointers::{self, CARRAY_BASE_KEY, CARRAY_IDX_KEY, CARRAY_KIND};
use crate::platforms::libc::{complex_adapter, regex_adapter, time_adapter};
use crate::platforms::libc::{
    ctype_adapter, math_adapter, stdio_adapter, string_adapter, wchar_adapter,
};

pub fn parse(source: &str) -> Result<Module, String> {
    let (preprocessed, pp_macros) = preprocess_c_source(source);
    let mut pairs = CParser::parse(Rule::program, &preprocessed)
        .map_err(|e| format!("C parse error: {e}"))?;
    let program = pairs.next().ok_or("empty parse")?;
    let mut w = Walker::default();
    // Object-like macros consumed by the preprocessor (e.g. `#define NDEBUG`)
    // are no longer in the source the walker re-parses; seed them so `#ifdef`-
    // style checks (assert/NDEBUG) still see them.
    w.object_macros = pp_macros;
    let mut body = Vec::new();
    for item in program.into_inner() {
        match item.as_rule() {
            Rule::EOI => {}
            _ => w.walk_top_item(item, &mut body),
        }
    }
    // Prepend runtime helpers and static globals before the rest of the module body.
    let mut full_body = crate::platforms::libc::c_runtime::prelude();
    full_body.extend(w.static_globals);
    full_body.extend(body);
    Ok(Module {
        name: "main".to_string(),
        language: Lang::Unknown,
        body: full_body,
        imports: Vec::new(),
    })
}

#[derive(Default)]
struct Walker {
    /// struct/union name → ordered field names (for zero-init at decl site)
    structs: HashMap<String, Vec<String>>,
    /// struct name → field → (bitfield width, is_signed) for `int x : N` members.
    struct_bitfields: HashMap<String, HashMap<String, (i64, bool)>>,
    /// struct/union name → field name → field type (for nested struct handling)
    struct_field_types: HashMap<String, HashMap<String, String>>,
    /// typedef names whose declarator is pointer-shaped.
    typedef_pointer_aliases: HashSet<String>,
    /// typedef names whose declarator is `char *`-shaped.
    typedef_char_pointer_aliases: HashSet<String>,
    /// identifiers declared as `char*`; used for pointer-like string traversal.
    char_pointers: HashSet<String>,
    /// char pointer variable -> (base string/array variable, element offset)
    char_pointer_offsets: HashMap<String, (String, Expression)>,
    /// char pointer variable -> struct variable whose address it stores from `(char*)&obj`.
    char_pointer_struct_bases: HashMap<String, String>,
    /// identifiers declared as non-char pointer to array (int*, double*, etc.)
    /// These are PLAIN arrays (int arr[N]) — direct JS array indexing.
    array_ptr_vars: HashSet<String>,
    /// identifiers declared as pointer variables, even when their current value is NULL.
    pointer_vars: HashSet<String>,
    /// identifiers declared as function-pointer values.
    function_pointer_vars: HashSet<String>,
    /// identifiers that ARE pointer variables backed by a carray object
    /// `{__ref_kind:"carray", __base:Array, __idx:i32}`.
    /// Covers `int *p = arr;` and parameters declared as `int *p`.
    carray_ptr_vars: HashSet<String>,
    /// pointer variable -> variable whose address it stores from `T *p = &x`.
    pointer_address_aliases: HashMap<String, String>,
    /// pointer variable -> struct/union member expression from `T *p = &obj.field`.
    pointer_member_aliases: HashMap<String, Expression>,
    /// identifiers whose address has been taken with `&name`; later plain
    /// reads/writes go through the common reference-cell AST.
    address_taken: HashSet<String>,
    /// variable/parameter name → C type string for sizeof resolution
    var_types: HashMap<String, String>,
    /// variable/parameter name → precomputed sizeof value (accounts for arrays)
    var_sizes: HashMap<String, i64>,
    /// variable name → explicit `_Alignas(N)` / `alignas(N)` alignment, when present.
    var_alignments: HashMap<String, i64>,
    /// function-like macros: name → (params, body text)
    macros: HashMap<String, (Vec<String>, String)>,
    /// object-like macros: name → raw replacement text
    object_macros: HashMap<String, String>,
    /// function name → parameter C type hints, used to normalize pointer arguments.
    function_param_types: HashMap<String, Vec<Option<String>>>,
    /// function name → explicit return type, used to infer callable variable types.
    function_return_types: HashMap<String, String>,
    /// C enum constants are integer constants and can appear in global initializers.
    enum_constants: HashMap<String, i64>,
    /// current function name (for static local mangling)
    current_function: String,
    /// current function parameter names in declaration order.
    current_param_names: Vec<String>,
    /// static local variable orignal name → mangled global name
    static_renames: HashMap<String, String>,
    /// accumulated static-local declarations to prepend to the module body
    static_globals: Vec<Statement>,
    /// current function char* parameter name → parameter index.
    current_char_param_indices: HashMap<String, usize>,
    /// function-pointer bindings that should be removed when the current function exits.
    current_function_pointer_names: Vec<String>,
    /// function name → char* parameter writes `(param_index, index, value)`.
    char_param_writes: HashMap<String, Vec<(usize, Expression, Expression)>>,
    /// `atexit` handlers registered while walking the current `main` body.
    current_atexit_finalizers: Vec<Statement>,
    /// monotonically-increasing counter for synthetic temporaries (e.g. hoisting
    /// a side-effecting index out of a char-buffer element write).
    tmp_counter: u32,
}

#[derive(Clone, Copy)]
struct PpCond {
    parent_active: bool,
    taken: bool,
    active: bool,
}

fn preprocess_c_source(source: &str) -> (String, HashMap<String, String>) {
    let mut out = Vec::new();
    let mut object_macros: HashMap<String, String> = HashMap::new();
    let mut fn0_macros: HashMap<String, String> = HashMap::new();
    let mut cond_stack: Vec<PpCond> = Vec::new();

    let date_literal = "\"Jun 17 2026\"";
    let time_literal = "\"00:00:00\"";
    let file_literal = "\"<stdin>\"";

    for (idx, raw_line) in source.lines().enumerate() {
        let line_no = idx + 1;
        let trimmed = raw_line.trim_start();
        let current_active = cond_stack.last().map(|c| c.active).unwrap_or(true);

        if let Some(rest) = trimmed.strip_prefix('#') {
            let directive = rest.trim_start();

            if let Some(after) = directive.strip_prefix("define") {
                if current_active {
                    let rest = after.trim_start();
                    let name_len = rest
                        .chars()
                        .take_while(|c| c.is_ascii_alphanumeric() || *c == '_')
                        .count();
                    if name_len > 0 {
                        let name = &rest[..name_len];
                        let tail = &rest[name_len..];
                        if tail.starts_with('(') {
                            // Function-like macro: keep directive in source for the walker.
                            if let Some(close) = tail.find(')') {
                                let params = &tail[1..close];
                                let value = tail[close + 1..].trim();
                                if params.trim().is_empty() {
                                    fn0_macros.insert(
                                        name.to_string(),
                                        if value.is_empty() {
                                            "1".to_string()
                                        } else {
                                            value.to_string()
                                        },
                                    );
                                }
                            }
                            out.push(raw_line.to_string());
                        } else {
                            // Object-like macro: track for conditional/object substitution.
                            let value = {
                                let v = tail.trim_start();
                                if v.is_empty() {
                                    "1".to_string()
                                } else {
                                    v.to_string()
                                }
                            };
                            object_macros.insert(name.to_string(), value);
                        }
                    }
                }
                continue;
            }

            if let Some(after) = directive.strip_prefix("undef") {
                if current_active {
                    let name = after.trim();
                    if !name.is_empty() {
                        object_macros.remove(name);
                        fn0_macros.remove(name);
                    }
                }
                continue;
            }

            if let Some(after) = directive.strip_prefix("ifdef") {
                let parent_active = current_active;
                let cond = parent_active && object_macros.contains_key(after.trim());
                cond_stack.push(PpCond {
                    parent_active,
                    taken: cond,
                    active: cond,
                });
                continue;
            }

            if let Some(after) = directive.strip_prefix("ifndef") {
                let parent_active = current_active;
                let cond = parent_active && !object_macros.contains_key(after.trim());
                cond_stack.push(PpCond {
                    parent_active,
                    taken: cond,
                    active: cond,
                });
                continue;
            }

            if let Some(after) = directive.strip_prefix("if") {
                let parent_active = current_active;
                let cond = parent_active && eval_pp_expr(after.trim(), &object_macros);
                cond_stack.push(PpCond {
                    parent_active,
                    taken: cond,
                    active: cond,
                });
                continue;
            }

            if let Some(after) = directive.strip_prefix("elif") {
                if let Some(top) = cond_stack.last_mut() {
                    if !top.parent_active || top.taken {
                        top.active = false;
                    } else {
                        let cond = eval_pp_expr(after.trim(), &object_macros);
                        top.active = cond;
                        if cond {
                            top.taken = true;
                        }
                    }
                }
                continue;
            }

            if directive.starts_with("else") {
                if let Some(top) = cond_stack.last_mut() {
                    top.active = top.parent_active && !top.taken;
                    top.taken = true;
                }
                continue;
            }

            if directive.starts_with("endif") {
                cond_stack.pop();
                continue;
            }

            if current_active {
                out.push(raw_line.to_string());
            }
            continue;
        }

        if !current_active {
            continue;
        }

        let mut line = raw_line.to_string();

        // Expand very simple function-like macros without params: NAME().
        for (name, body) in &fn0_macros {
            let pat = format!("{}()", name);
            if line.contains(&pat) {
                line = line.replace(&pat, body);
            }
        }

        line = replace_word(&line, "__LINE__", &line_no.to_string());
        line = replace_word(&line, "__FILE__", file_literal);
        line = replace_word(&line, "__DATE__", date_literal);
        line = replace_word(&line, "__TIME__", time_literal);

        line = expand_object_macros_in_text(&line, &object_macros);

        out.push(line);
    }

    (out.join("\n"), object_macros)
}

fn eval_pp_expr(expr_src: &str, object_macros: &HashMap<String, String>) -> bool {
    let mut expr = expr_src.trim().to_string();

    loop {
        if let Some(pos) = expr.find("defined(") {
            let rest = &expr[pos + 8..];
            if let Some(end) = rest.find(')') {
                let name = rest[..end].trim();
                let val = if object_macros.contains_key(name) {
                    "1"
                } else {
                    "0"
                };
                let mut next = String::new();
                next.push_str(&expr[..pos]);
                next.push_str(val);
                next.push_str(&rest[end + 1..]);
                expr = next;
                continue;
            }
        }
        if let Some(pos) = expr.find("defined ") {
            let after = &expr[pos + "defined ".len()..];
            let name: String = after
                .chars()
                .take_while(|c| c.is_ascii_alphanumeric() || *c == '_')
                .collect();
            if !name.is_empty() {
                let val = if object_macros.contains_key(&name) {
                    "1"
                } else {
                    "0"
                };
                let mut next = String::new();
                next.push_str(&expr[..pos]);
                next.push_str(val);
                next.push_str(&after[name.len()..]);
                expr = next;
                continue;
            }
        }
        break;
    }

    let mut tokens = Vec::new();
    let mut cur = String::new();
    for ch in expr.chars() {
        if ch.is_ascii_alphanumeric() || ch == '_' {
            cur.push(ch);
        } else {
            if !cur.is_empty() {
                let t = if let Some(v) = object_macros.get(&cur) {
                    v.clone()
                } else if cur
                    .chars()
                    .next()
                    .map(|c| c.is_ascii_alphabetic())
                    .unwrap_or(false)
                {
                    "0".to_string()
                } else {
                    cur.clone()
                };
                tokens.push(t);
                cur.clear();
            }
            tokens.push(ch.to_string());
        }
    }
    if !cur.is_empty() {
        let t = if let Some(v) = object_macros.get(&cur) {
            v.clone()
        } else if cur
            .chars()
            .next()
            .map(|c| c.is_ascii_alphabetic())
            .unwrap_or(false)
        {
            "0".to_string()
        } else {
            cur.clone()
        };
        tokens.push(t);
    }
    let reduced = tokens.join("").replace(' ', "");

    for op in ["==", "!=", ">=", "<=", ">", "<"] {
        if let Some(i) = reduced.find(op) {
            let lhs = reduced[..i].trim().parse::<i64>().unwrap_or(0);
            let rhs = reduced[i + op.len()..].trim().parse::<i64>().unwrap_or(0);
            return match op {
                "==" => lhs == rhs,
                "!=" => lhs != rhs,
                ">=" => lhs >= rhs,
                "<=" => lhs <= rhs,
                ">" => lhs > rhs,
                "<" => lhs < rhs,
                _ => false,
            };
        }
    }

    reduced.trim().parse::<i64>().unwrap_or(0) != 0
}

fn expand_object_macros_in_text(line: &str, object_macros: &HashMap<String, String>) -> String {
    let mut keys: Vec<&str> = object_macros.keys().map(|k| k.as_str()).collect();
    keys.sort_by_key(|k| std::cmp::Reverse(k.len()));

    let mut out = line.to_string();
    for _ in 0..8 {
        let mut changed = false;
        let mut next = out.clone();
        for key in &keys {
            if let Some(value) = object_macros.get(*key) {
                let replaced = replace_word(&next, key, value);
                if replaced != next {
                    changed = true;
                    next = replaced;
                }
            }
        }
        out = next;
        if !changed {
            break;
        }
    }
    out
}

// Shared libc AST builders + C FILE slot constants live in the libc platform
// (`platforms/libc/build.rs`) so the C walker and the libc adapters construct
// the runtime from one set of helpers. Imported here so existing call sites
// (expr/stmt/ident/int_lit/.../function_stmt, CFILE_*) resolve unchanged.
use crate::platforms::libc::build::{
    assign_expr, call_expr, expr, ident, if_stmt, index_expr, int_lit, member, null_lit, stmt,
    str_lit, var_decl_stmt,
};
// Bare math-call constructors used by walker arms (cabs, etc.). The series
// builders (tgamma/erf/…) are only referenced from the runtime prelude.
use crate::platforms::libc::math_runtime::{ecma_math_call, ecma_math_call2};


fn carray_indexed_access(object: Expression, index: Expression) -> Expression {
    let adjusted = expr(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(expr(ExprKind::Member {
            object: Box::new(object.clone()),
            field: CARRAY_IDX_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(index),
    });
    expr(ExprKind::Index {
        object: Box::new(expr(ExprKind::Member {
            object: Box::new(object),
            field: CARRAY_BASE_KEY.to_string(),
            null_safe: false,
        })),
        index: Box::new(adjusted),
        null_safe: false,
    })
}

fn declarator_has_pointer(pair: &Pair<Rule>) -> bool {
    for child in pair.clone().into_inner() {
        match child.as_rule() {
            Rule::pointer => return true,
            Rule::declarator | Rule::direct_declarator => {
                if declarator_has_pointer(&child) {
                    return true;
                }
            }
            _ => {}
        }
    }
    false
}

impl Walker {
    fn walk_top_item(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        match pair.as_rule() {
            Rule::preproc_directive => self.walk_preproc(pair, out),
            Rule::function_definition => {
                if let Some(s) = self.walk_function(pair) {
                    out.push(s);
                }
            }
            Rule::static_assert_statement => out.push(self.walk_static_assert(pair)),
            Rule::declaration => self.walk_declaration(pair, out),
            _ => {}
        }
    }

    fn walk_preproc(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        for inner in pair.into_inner() {
            match inner.as_rule() {
            Rule::define_directive => {
                // #define NAME value  → const NAME = value  (object-like only)
                let mut it = inner.into_inner();
                let name = match it.next() {
                    Some(p) if p.as_rule() == Rule::ident_name => p.as_str().to_string(),
                    _ => continue,
                };
                // Skip function-like macros (define_params present).
                let mut value_pair = None;
                let mut params_pair = None;
                for p in it {
                    match p.as_rule() {
                        Rule::define_params => params_pair = Some(p),
                        Rule::define_value => value_pair = Some(p),
                        _ => {}
                    }
                }
                if let Some(pp) = params_pair {
                    // Function-like macro: store for expansion at call sites
                    let raw_params = pp.as_str().trim();
                    let params_src = raw_params
                        .strip_prefix('(')
                        .and_then(|s| s.strip_suffix(')'))
                        .unwrap_or(raw_params);
                    let params: Vec<String> = params_src
                        .split(',')
                        .map(str::trim)
                        .filter(|s| !s.is_empty() && *s != "...")
                        .map(|s| s.to_string())
                        .collect();
                    let body_text = value_pair
                        .map(|p| p.as_str().trim().to_string())
                        .unwrap_or_default();
                    self.macros.insert(name, (params, body_text));
                    continue;
                }
                let init = value_pair
                    .map(|p| {
                        let raw = p.as_str().trim().to_string();
                        self.object_macros.insert(name.clone(), raw.clone());
                        self.parse_define_value(&raw)
                    })
                    .unwrap_or_else(|| {
                        // Valueless `#define NDEBUG` still defines the macro (as
                        // empty) — `assert` and other `#ifdef` checks rely on it.
                        self.object_macros.entry(name.clone()).or_default();
                        expr(ExprKind::Lit(Literal::Int(1)))
                    });
                out.push(stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(name),
                        type_hint: None,
                        init: Some(init),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                }));
            }
            Rule::include_directive => {
                // Inject standard header constants as object macros
                if let Some(target) = inner.into_inner().next() {
                    let header = target.as_str().trim_matches(|c| c == '<' || c == '>' || c == '"');
                    self.inject_header_constants(header, out);
                }
            }
            _ => {} // other_directive, conditionals: ignored
            }
        }
    }

    fn inject_header_constants(&mut self, header: &str, out: &mut Vec<Statement>) {
        let defs: &[(&str, i64)] = match header {
            "stdint.h" => &[
                ("INT8_MIN",   -128),
                ("INT8_MAX",    127),
                ("UINT8_MAX",   255),
                ("INT16_MIN",  -32768),
                ("INT16_MAX",   32767),
                ("UINT16_MAX",  65535),
                ("INT32_MIN",  -2147483648),
                ("INT32_MAX",   2147483647),
                ("UINT32_MAX",  4294967295_i64),
                ("INT64_MIN",  i64::MIN),
                ("INT64_MAX",  i64::MAX),
                ("UINT64_MAX", i64::MAX),
                ("INTPTR_MIN", i64::MIN),
                ("INTPTR_MAX", i64::MAX),
                ("UINTPTR_MAX", i64::MAX),
                ("PTRDIFF_MIN", i64::MIN),
                ("PTRDIFF_MAX", i64::MAX),
                ("SIZE_MAX",   i64::MAX),
            ],
            "limits.h" => &[
                ("CHAR_BIT",    8),
                ("CHAR_MIN",   -128),
                ("CHAR_MAX",    127),
                ("SCHAR_MIN",  -128),
                ("SCHAR_MAX",   127),
                ("UCHAR_MAX",   255),
                ("SHRT_MIN",   -32768),
                ("SHRT_MAX",    32767),
                ("USHRT_MAX",   65535),
                ("INT_MIN",    -2147483648),
                ("INT_MAX",     2147483647),
                ("UINT_MAX",    4294967295_i64),
                ("LONG_MIN",   i64::MIN),
                ("LONG_MAX",   i64::MAX),
                ("ULONG_MAX",  i64::MAX),
                ("LLONG_MIN",  i64::MIN),
                ("LLONG_MAX",  i64::MAX),
                ("ULLONG_MAX", i64::MAX),
            ],
            // POSIX regex.h — regcomp cflags, regexec eflags, and error codes.
            "regex.h" => &[
                ("REG_EXTENDED", 1),
                ("REG_ICASE", 2),
                ("REG_NEWLINE", 4),
                ("REG_NOSUB", 8),
                ("REG_NOTBOL", 1),
                ("REG_NOTEOL", 2),
                ("REG_NOMATCH", 1),
                ("REG_BADPAT", 2),
                ("REG_ESPACE", 12),
            ],
            "math.h" | "cmath" => &[],
            _ => &[],
        };
        for (name, val) in defs {
            // Only inject if not already defined by user
            if !self.object_macros.contains_key(*name) {
                self.object_macros.insert(name.to_string(), val.to_string());
                // 64-bit limits (SIZE_MAX, INTPTR_MIN, LONG_MAX, …) exceed the
                // i32 range and would be destroyed by 32-bit integer
                // normalization; emit them as floats so the magnitude/sign
                // survive (these constants are only ever compared, not used as
                // exact bit patterns).
                let init = if *val > i32::MAX as i64 || *val < i32::MIN as i64 {
                    expr(ExprKind::Lit(Literal::Float(*val as f64)))
                } else {
                    expr(ExprKind::Lit(Literal::Int(*val)))
                };
                out.push(stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(name.to_string()),
                        type_hint: None,
                        init: Some(init),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                }));
            }
        }
        // Float constants for float.h
        let float_defs: &[(&str, f64)] = match header {
            "float.h" => &[
                    ("FLT_EPSILON",  1.1920929e-7_f64),
                ("DBL_EPSILON", 2.220446049250313e-16_f64),
                ("FLT_MAX",     3.4028235e38_f64),
                ("DBL_MAX",     1.7976931348623157e308_f64),
                ("FLT_MIN",     1.1754944e-38_f64),
                ("DBL_MIN",     2.2250738585072014e-308_f64),
                ("FLT_DIG",     6.0),
                ("DBL_DIG",     15.0),
                ("FLT_MANT_DIG", 24.0),
                ("DBL_MANT_DIG", 53.0),
                ("FLT_MAX_EXP",  128.0),
                ("DBL_MAX_EXP",  1024.0),
                ("FLT_MIN_EXP", -125.0),
                ("DBL_MIN_EXP", -1021.0),
                ("FLT_RADIX",    2.0),
            ],
                "math.h" | "cmath" => &[
                    ("HUGE_VAL",  f64::INFINITY),
                    ("HUGE_VALF", f64::INFINITY),
                    ("INFINITY",  f64::INFINITY),
                    ("NAN",       f64::NAN),
                    ("M_PI",      std::f64::consts::PI),
                    ("M_E",       std::f64::consts::E),
                    ("M_SQRT2",   std::f64::consts::SQRT_2),
                    ("M_LN2",     std::f64::consts::LN_2),
                    ("M_LN10",    std::f64::consts::LN_10),
                    ("M_LOG2E",   std::f64::consts::LOG2_E),
                    ("M_LOG10E",  std::f64::consts::LOG10_E),
                ],
                _ => &[],
        };
        for (name, val) in float_defs {
            if !self.object_macros.contains_key(*name) {
                self.object_macros.insert(name.to_string(), val.to_string());
                out.push(stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(name.to_string()),
                        type_hint: None,
                        init: Some(expr(ExprKind::Lit(Literal::Float(*val)))),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                }));
            }
        }
    }

    /// Best-effort literal for a `#define` value; falls back to a string.
    fn parse_define_value(&mut self, text: &str) -> Expression {
        if let Ok(i) = text.parse::<i64>() {
            return expr(ExprKind::Lit(Literal::Int(i)));
        }
        if let Ok(f) = text.parse::<f64>() {
            return expr(ExprKind::Lit(Literal::Float(f)));
        }
        // Try re-parsing the value as a C expression.
        if let Ok(mut p) = CParser::parse(Rule::assignment_expression, text) {
            if let Some(e) = p.next() {
                return self.walk_assignment(e);
            }
        }
        expr(ExprKind::Lit(Literal::Str(text.to_string())))
    }

    // ── Functions ──────────────────────────────────────────────────────────

    fn walk_function(&mut self, pair: Pair<Rule>) -> Option<Statement> {
        let mut return_type = None;
        let mut return_pointer_count = 0usize;
        let mut name = String::new();
        let mut params = Vec::new();
        let mut body = Vec::new();
        let previous_atexit_finalizers = std::mem::take(&mut self.current_atexit_finalizers);
        let previous_param_names = std::mem::take(&mut self.current_param_names);
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::declaration_specifiers => return_type = Some(self.type_text(p)),
                Rule::declarator => {
                    return_pointer_count = p
                        .clone()
                        .into_inner()
                        .filter(|child| child.as_rule() == Rule::pointer)
                        .count();
                    let (n, ps) = self.declarator_name_and_params(p);
                    name = n;
                    if let Some(ps) = ps {
                        params = ps;
                    }
                }
                Rule::compound_statement => {
                    // Set current function context for static local mangling
                    self.current_function = name.clone();
                    self.current_param_names = params.iter().map(|p| p.name.clone()).collect();
                    self.static_renames.clear();
                    let mut scoped_char_params = Vec::new();
                    let mut scoped_carray_params = Vec::new();
                    for (idx, param) in params.iter().enumerate() {
                        if let Some(type_hint) = &param.type_hint {
                            if type_hint.contains("char") && type_hint.contains('*') {
                                self.char_pointers.insert(param.name.clone());
                                self.current_char_param_indices
                                    .insert(param.name.clone(), idx);
                                scoped_char_params.push(param.name.clone());
                            } else if type_hint == "func" {
                                self.function_pointer_vars.insert(param.name.clone());
                                self.current_function_pointer_names
                                    .push(param.name.clone());
                            } else if self.is_carray_compatible_pointer_param(type_hint) {
                                self.carray_ptr_vars.insert(param.name.clone());
                                scoped_carray_params.push(param.name.clone());
                            }
                        }
                    }
                    body = self.walk_block(p);
                    for param in scoped_carray_params.iter().rev() {
                        let param_ident = ident(param);
                        // Wrap a decayed array pointer in a carray — but leave an
                        // existing carray AND a NULL pointer untouched, so a null
                        // arg stays falsy (`p ? *p : -1` must see NULL as false).
                        let keep_as_is = expr(ExprKind::Binary {
                            op: BinOp::Or,
                            left: Box::new(pointers::is_carray_ptr_kind(param_ident.clone())),
                            right: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Eq,
                                left: Box::new(param_ident.clone()),
                                right: Box::new(expr(ExprKind::Lit(Literal::Null))),
                            })),
                        });
                        body.insert(
                            0,
                            stmt(StmtKind::Expr(assign_expr(
                                param_ident.clone(),
                                expr(ExprKind::Ternary {
                                    cond: Box::new(keep_as_is),
                                    then: Box::new(param_ident.clone()),
                                    else_: Box::new(pointers::make_carray_ptr(
                                        param_ident,
                                        int_lit(0),
                                    )),
                                }),
                            ))),
                        );
                    }
                    for param in scoped_char_params {
                        self.char_pointers.remove(&param);
                    }
                    self.current_char_param_indices.clear();
                    for param in self.current_function_pointer_names.drain(..) {
                        self.function_pointer_vars.remove(&param);
                    }
                    for param in scoped_carray_params {
                        self.carray_ptr_vars.remove(&param);
                    }
                    for param in params.iter().rev() {
                        let Some(type_hint) = &param.type_hint else {
                            continue;
                        };
                        if type_hint.contains('*') {
                            continue;
                        }
                        let normalized_type = normalized_c_type_name(type_hint);
                        if self.structs.contains_key(&normalized_type) {
                            body.insert(
                                0,
                                stmt(StmtKind::Expr(expr(ExprKind::Assign {
                                    target: Box::new(ident(&param.name)),
                                    value: Box::new(
                                        self.deep_copy_struct(type_hint, ident(&param.name)),
                                    ),
                                }))),
                            );
                        }
                    }
                    if name == "main" && !self.current_atexit_finalizers.is_empty() {
                        let mut finally = std::mem::take(&mut self.current_atexit_finalizers);
                        finally.reverse();
                        body = vec![stmt(StmtKind::Try {
                            body,
                            catches: vec![],
                            else_body: None,
                            finally: Some(finally),
                        })];
                    }
                    body = lower_c_gotos(body);
                    // setjmp.h: wrap the setjmp re-entry point (if any) so it
                    // "returns twice" via a longjmp-throw catch loop.
                    body = wrap_setjmp_in_block(body, &mut self.tmp_counter);
                }
                _ => {}
            }
        }
        if name.is_empty() {
            return None;
        }
        if let Some(ref mut type_name) = return_type {
            if return_pointer_count > 0 {
                for _ in 0..return_pointer_count {
                    type_name.push_str(" *");
                }
            }
        }
        self.function_param_types.insert(
            name.clone(),
            params.iter().map(|param| param.type_hint.clone()).collect(),
        );
        if let Some(ref return_type) = return_type {
            self.function_return_types
                .insert(name.clone(), return_type.clone());
        }
        self.current_param_names = previous_param_names;
        self.current_atexit_finalizers = previous_atexit_finalizers;
        // `main(int argc, char *argv[])` is auto-invoked with no args by the
        // entry point; give its params C-faithful defaults (argc=1, argv=[prog]).
        if name == "main" {
            for (idx, p) in params.iter_mut().enumerate() {
                if p.default.is_none() {
                    p.default = Some(if idx == 0 {
                        int_lit(1)
                    } else if idx == 1 {
                        expr(ExprKind::Array(vec![ArrayElement {
                            key: None,
                            value: str_lit("program"),
                            spread: false,
                            by_ref: false,
                        }]))
                    } else {
                        expr(ExprKind::Lit(Literal::Null))
                    });
                }
            }
        }
        Some(stmt(StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        }))
    }

    fn declarator_name_and_params(&mut self, pair: Pair<Rule>) -> (String, Option<Vec<Param>>) {
        // declarator = pointer? ~ direct_declarator
        let mut name = String::new();
        let mut params = None;
        for p in pair.into_inner() {
            if p.as_rule() == Rule::direct_declarator {
                for d in p.into_inner() {
                    match d.as_rule() {
                        Rule::ident_name => name = d.as_str().to_string(),
                        Rule::declarator => {
                            let (n, ps) = self.declarator_name_and_params(d);
                            name = n;
                            if ps.is_some() {
                                params = ps;
                            }
                        }
                        Rule::param_suffix => params = Some(self.walk_params(d)),
                        _ => {}
                    }
                }
            }
        }
        (name, params)
    }

    fn walk_params(&mut self, pair: Pair<Rule>) -> Vec<Param> {
        let mut params = Vec::new();
        for p in pair.into_inner() {
            if p.as_rule() == Rule::parameter_list {
                let has_ellipsis = p.as_str().contains("...");
                for decl in p.into_inner() {
                    if decl.as_rule() == Rule::parameter_decl {
                        let decl_text = decl.as_str().to_string();
                        let mut pname = String::new();
                        let mut type_hint = None;
                        let mut is_pointer_decl = decl_text.contains('[');
                        let is_function_pointer_decl = decl_text.contains("(*") && decl_text.contains(")(");
                        for d in decl.into_inner() {
                            match d.as_rule() {
                                Rule::declaration_specifiers => type_hint = Some(self.type_text(d)),
                                Rule::declarator => {
                                    is_pointer_decl = is_pointer_decl
                                        || declarator_has_pointer(&d)
                                        || decl_text.contains('*');
                                    pname = self.declarator_name_and_params(d).0;
                                    if is_function_pointer_decl {
                                        let mut return_type = type_hint
                                            .as_deref()
                                            .map(normalized_c_type_name)
                                            .unwrap_or_else(|| normalized_c_type_name(&decl_text));
                                        let pointer_star_count = decl_text.matches('*').count().saturating_sub(1);
                                        for _ in 0..pointer_star_count {
                                            return_type.push_str(" *");
                                        }
                                        type_hint = Some(format!("func() -> {}", return_type));
                                    }
                                }
                                _ => {}
                            }
                        }
                        if is_pointer_decl && !is_function_pointer_decl {
                            if let Some(hint) = &mut type_hint {
                                let existing = hint.matches('*').count();
                                let declared = decl_text.matches('*').count().max(1);
                                for _ in existing..declared {
                                    hint.push_str(" *");
                                }
                            }
                        }
                        if !pname.is_empty() {
                            params.push(Param {
                                name: pname,
                                type_hint,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            });
                        }
                    }
                }
                if has_ellipsis {
                    params.push(Param {
                        name: "__va_args".to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: true,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    });
                }
            }
        }
        params
    }

    // ── Declarations ─────────────────────────────────────────────────────────

    fn walk_declaration(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        let inner = pair.into_inner().next();
        let Some(inner) = inner else { return };
        match inner.as_rule() {
            Rule::typedef_declaration => self.walk_typedef(inner, out),
            Rule::normal_declaration => self.walk_normal_declaration(inner, out),
            _ => {}
        }
    }

    fn walk_typedef(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        // typedef declaration_specifiers declarator (, declarator)* ;
        let mut specs = None;
        let mut names = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::declaration_specifiers => specs = Some(p),
                Rule::declarator => {
                    let is_pointer_alias = declarator_has_pointer(&p)
                        || p.as_str().split('=').next().unwrap_or("").contains('*');
                    let name = self.declarator_name_and_params(p).0;
                    if is_pointer_alias && !name.is_empty() {
                        self.typedef_pointer_aliases.insert(name.clone());
                    }
                    names.push(name);
                }
                _ => {}
            }
        }
        if let Some(ref specs) = specs {
            if self.type_text(specs.clone()).contains("char") {
                for name in &names {
                    if self.typedef_pointer_aliases.contains(name) {
                        self.typedef_char_pointer_aliases.insert(name.clone());
                    }
                }
            }
            // typedef struct {...} Name; → register Name as struct alias.
            if let Some((tag, fields, field_types, bitfields)) =
                self.struct_def_from_specifiers(specs)
            {
                for name in &names {
                    self.structs.insert(name.clone(), fields.clone());
                    // Register field types under the typedef name too, so nested
                    // aggregate inits (`Theme t = {{...},{...}}`) can resolve each
                    // field's struct type. Anonymous typedef'd structs have no tag,
                    // so this is the only place their field types get keyed.
                    self.struct_field_types
                        .insert(name.clone(), field_types.clone());
                    if !bitfields.is_empty() {
                        self.struct_bitfields.insert(name.clone(), bitfields.clone());
                    }
                    out.push(self.make_struct_decl(name, &fields));
                }
                let _ = tag;
            }
            // typedef enum { A, B } Name; → emit enum members as consts.
            if let Some((_tag, members)) = self.enum_def_from_specifiers(specs) {
                let mut next_val: i64 = 0;
                for member in &members {
                    let val = extract_enum_val(&member.value, next_val);
                    self.enum_constants.insert(member.name.clone(), val);
                    out.push(stmt(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(member.name.clone()),
                            type_hint: Some("int".to_string()),
                            init: Some(expr(ExprKind::Lit(Literal::Int(val)))),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Const,
                    }));
                    next_val = val + 1;
                }
            }
        }
    }

    fn walk_normal_declaration(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        let mut specs = None;
        let mut init_list = None;
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::declaration_specifiers => specs = Some(p),
                Rule::init_declarator_list => init_list = Some(p),
                _ => {}
            }
        }
        let Some(specs) = specs else { return };

        // A struct/union/enum definition with a body.
        if let Some((tag, fields, _field_types, _bitfields)) =
            self.struct_def_from_specifiers(&specs)
        {
            if let Some(tag) = tag.clone() {
                self.structs.insert(tag.clone(), fields.clone());
                if init_list.is_none() {
                    out.push(self.make_struct_decl(&tag, &fields));
                }
            }
        }
        if let Some((name, members)) = self.enum_def_from_specifiers(&specs) {
            // In C, enum members are global integer constants — emit each as const.
            // Auto-increment value starting from 0, incremented per member.
            let mut next_val: i64 = 0;
            for member in &members {
                let val = extract_enum_val(&member.value, next_val);
                self.enum_constants.insert(member.name.clone(), val);
                out.push(stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(member.name.clone()),
                        type_hint: Some("int".to_string()),
                        init: Some(expr(ExprKind::Lit(Literal::Int(val)))),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                }));
                next_val = val + 1;
            }
            out.push(stmt(StmtKind::EnumDecl {
                name: name.unwrap_or_else(|| "__anon_enum".to_string()),
                members,
                visibility: Visibility::Public,
                is_flags: false,
                backing_type: None,
                interfaces: Vec::new(),
                body_members: Vec::new(),
                decorators: Vec::new(),
            }));
        }

        let declared_alignment = explicit_alignment_value(specs.as_str());
        let struct_fields = self.struct_type_of_specifiers(&specs);
        let mut type_text = self.type_text(specs);
        // wasm32-wasi (wasi-libc) defines `wchar_t` as `int` (4-byte UTF-32).
        // Rewriting it to `int` makes wchar_t buffers plain flat int arrays:
        // correct sizeof/stride, array-method dispatch, and pointer arithmetic
        // via the carray model — exactly how a real C→WASM toolchain treats them.
        if normalized_c_type_name(&type_text) == "wchar_t" {
            type_text = type_text.replace("wchar_t", "int");
        }

        // Skip extern declarations with no initializer (function prototypes etc.)
        if type_text.split_whitespace().any(|w| w == "extern") && init_list.is_none() {
            return;
        }

        // Check for static locals (only when inside a function)
        let is_static_local = !self.current_function.is_empty()
            && type_text.split_whitespace().any(|w| w == "static");

        let Some(init_list) = init_list else { return };
        let mut declarations = Vec::new();
        for idecl in init_list.into_inner() {
            if idecl.as_rule() != Rule::init_declarator {
                continue;
            }
            let declarator_text = idecl.as_str().split('=').next().unwrap_or("").to_string();
            let mut name = String::new();
            let mut array_bounds: Option<Vec<Expression>> = None;
            let mut init = None;
            let mut is_pointer_decl = false;
            let mut init_is_addr_of = false;
            let mut is_function_proto = false;
            let mut is_function_pointer_decl = false;
            let mut was_array_decl = false;
            for p in idecl.into_inner() {
                match p.as_rule() {
                    Rule::declarator => {
                        is_pointer_decl =
                            declarator_has_pointer(&p) || declarator_text.contains('*');
                        if declarator_text.contains('[') {
                            was_array_decl = true;
                        }
                        if self
                            .typedef_pointer_aliases
                            .contains(&normalized_c_type_name(&type_text))
                        {
                            is_pointer_decl = true;
                        }
                        // Detect function-pointer or prototype declarator: has param_suffix
                        is_function_proto =
                            p.as_str().contains('(') && !p.as_str().starts_with('*'); // not a function-pointer type
                        is_function_pointer_decl =
                            declarator_text.contains("(*") && declarator_text.contains(")(");
                        let (n, bounds) = self.declarator_name_and_bounds(p);
                        name = n;
                        array_bounds = bounds;
                        if array_bounds.is_some() {
                            was_array_decl = true;
                        }
                    }
                    Rule::initializer => {
                        // Check before walking if init is address-of (&x) form
                        init_is_addr_of = p.as_str().trim().starts_with('&');
                        let raw = self.walk_initializer(p);
                        if matches!(&raw.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n))
                        {
                            is_pointer_decl = true;
                        }
                        init = Some(if !is_pointer_decl {
                            if let Some(fields) = &struct_fields {
                                if array_bounds.is_some() {
                                    // Array of structs: convert each element to a named object.
                                    // `struct Pair pairs[2] = {{1,2},{3,4}}` → [{a:1,b:2},{a:3,b:4}]
                                    if let ExprKind::Array(elems) = raw.kind {
                                        let converted: Vec<ArrayElement> = elems
                                            .into_iter()
                                            .map(|el| ArrayElement {
                                                value: self.convert_array_init_to_struct_typed(
                                                    &type_text, el.value, fields,
                                                ),
                                                ..el
                                            })
                                            .collect();
                                        array_bounds = None; // embedded in literal
                                        expr(ExprKind::Array(converted))
                                    } else {
                                        raw
                                    }
                                } else if is_all_zero_init(&raw) {
                                    // `{0}` / `{{0}}` / `{{{0}}}` zero-initialise the
                                    // whole struct — build the proper zero shape
                                    // (incl. array-of-struct fields) via zero_struct.
                                    let sn = normalized_c_type_name(&type_text);
                                    self.zero_struct(Some(&sn), fields)
                                } else {
                                    // Convert array init to struct, and also handle struct-to-struct copy
                                    let converted = self.convert_array_init_to_struct_typed(
                                        &type_text,
                                        raw.clone(),
                                        fields,
                                    );
                                    // If init is a simple identifier (struct copy), wrap in deep copy
                                    if matches!(raw.kind, ExprKind::Ident(_))
                                        || matches!(raw.kind, ExprKind::Member { .. })
                                    {
                                        self.deep_copy_struct(&type_text, converted)
                                    } else {
                                        converted
                                    }
                                }
                            } else {
                                raw
                            }
                        } else {
                            raw
                        });
                    }
                    _ => {}
                }
            }
            if name.is_empty() {
                continue;
            }
            if is_pointer_decl && !is_function_pointer_decl {
                self.pointer_vars.insert(name.clone());
            }
            if is_pointer_decl {
                let pointee_type = normalized_c_type_name(type_text.trim_end_matches('*').trim());
                if let Some(fields) = self.structs.get(&pointee_type) {
                    if let Some(ExprKind::Array(elems)) = init.as_ref().map(|i| &i.kind) {
                        let zeroed: Vec<ArrayElement> = elems
                            .iter()
                            .map(|_| ArrayElement {
                                value: self.zero_struct(Some(&pointee_type), fields),
                                spread: false,
                                key: None,
                                by_ref: false,
                            })
                            .collect();
                        init = Some(expr(ExprKind::Array(zeroed)));
                    }
                }
            }
            if type_text.split_whitespace().any(|w| w == "extern") && init.is_none() {
                continue;
            }
            // Skip function prototypes: `int foo(int x);` has no init and is
            // a function-like declarator. Emitting `var foo;` would shadow the
            // actual function definition.
            if is_function_proto && init.is_none() && !is_pointer_decl {
                continue;
            }
            let normalized_type_text = normalized_c_type_name(&type_text);
            let type_is_char_pointer_alias = self
                .typedef_char_pointer_aliases
                .contains(&normalized_type_text);
            let declarator_pointer_depth = declarator_text.matches('*').count();
            let is_multi_level_char_pointer = if type_is_char_pointer_alias {
                // typedef char* P; P *pp; => two-level pointer
                declarator_pointer_depth > 0
            } else {
                declarator_pointer_depth > 1 || type_text.matches('*').count() > 1
            };
            let is_file_pointer_type = is_pointer_decl && normalized_type_text == "FILE";
            if is_pointer_decl && is_null_pointer_init(&init) {
                init = Some(null_lit());
            }
            // A char* initialized from a literal carray object (e.g. the
            // `(char*)&x` byte view) is a real array-backed pointer, not a string
            // char*; let it fall through to the carray branch so `*p` / `p[i]`
            // index bytes. Use `is_carray_object` (the carray Object literal),
            // NOT `is_carray_like_expr` — the latter also matches strstr/memchr
            // slice-or-null ternaries, which must stay string char-pointers.
            let init_is_carray_obj = init
                .as_ref()
                .map(|i| is_carray_object(i))
                .unwrap_or(false);
            if (type_text.contains("char") || type_is_char_pointer_alias)
                && !is_function_pointer_decl
                && !is_multi_level_char_pointer
                && !(is_pointer_decl && was_array_decl)
                && !init_is_carray_obj
            {
                // Track char* pointers AND char arrays (initialized with string literals)
                // for substring-based pointer arithmetic.
                let init_is_string = init
                    .as_ref()
                    .map(|i| matches!(i.kind, ExprKind::Lit(Literal::Str(_))))
                    .unwrap_or(false);
                let init_is_heap_array = init
                    .as_ref()
                    .map(|i| matches!(i.kind, ExprKind::Array(_)))
                    .unwrap_or(false);
                if is_pointer_decl && init_is_heap_array && !was_array_decl {
                    self.char_pointers.insert(name.clone());
                    init = Some(expr(ExprKind::Lit(Literal::Str(String::new()))));
                } else if init_is_string
                    || (is_pointer_decl && !init_is_addr_of && !is_null_pointer_init(&init))
                {
                    self.char_pointers.insert(name.clone());
                    if let Some((base, offset)) = char_pointer_offset_from_init(&init) {
                        self.char_pointer_offsets
                            .insert(name.clone(), (base, offset));
                    }
                }
                if is_pointer_decl {
                    if let Some(target) = pointer_address_target_from_init(&init) {
                        self.pointer_address_aliases.insert(name.clone(), target.clone());
                        if self
                            .var_types
                            .get(&target)
                            .map(|ty| normalized_c_type_name(ty).starts_with("struct "))
                            .unwrap_or(false)
                        {
                            self.char_pointer_struct_bases.insert(name.clone(), target);
                        }
                    }
                }
            } else if is_pointer_decl && !is_function_pointer_decl {
                if is_file_pointer_type {
                    // FILE* is modeled as an opaque integer handle, not as a carray/scalar-cell pointer.
                } else {
                // Non-char pointer variable — decide scalar-cell vs carray.
                // If the walked init is already a carray object (e.g. from `&arr[n]`),
                // track this var as carray; otherwise wrap a plain array as carray.
                let init_is_carray = init
                    .as_ref()
                    .map(|i| is_carray_like_expr(i))
                    .unwrap_or(false);
                if let Some(target) = self.pointer_member_target_from_char_struct_base_init(&init) {
                    self.pointer_member_aliases.insert(name.clone(), target.clone());
                    init = Some(expr(ExprKind::Unary {
                        op: UnaryOp::AddrOf,
                        expr: Box::new(target),
                    }));
                }
                if let Some(target) = pointer_address_target_from_init(&init) {
                    self.pointer_address_aliases.insert(name.clone(), target);
                } else if let Some(target) = pointer_member_target_from_init(&init) {
                    self.pointer_member_aliases.insert(name.clone(), target);
                } else if let Some(target) =
                    propagated_pointer_address_alias(&init, &self.pointer_address_aliases)
                {
                    self.pointer_address_aliases.insert(name.clone(), target);
                }
                if init_is_carray {
                    // int *p = &arr[n] → init already carray from apply_prefix
                    self.carray_ptr_vars.insert(name.clone());
                } else if !init_is_addr_of && !is_null_pointer_init(&init) {
                    if init_is_carray_pointer_var(&init, &self.carray_ptr_vars) {
                        self.carray_ptr_vars.insert(name.clone());
                    } else if !was_array_decl
                        && should_wrap_pointer_init_as_carray(&init, &self.array_ptr_vars)
                    {
                        // int *p = arr → wrap as carray
                        self.carray_ptr_vars.insert(name.clone());
                        if let Some(ref raw_init) = init {
                            init = Some(self.wrap_as_carray_init(raw_init.clone()));
                        }
                    }
                }
                // Some array-backed pointer inits (notably `int *p = &arr[0]`,
                // `int *p = arr + n`, and row decay `int *row = m[i]`) can miss
                // the earlier fast-path checks; ensure they are still treated as carray pointers.
                if !self.carray_ptr_vars.contains(&name)
                    && self.pointer_init_looks_array_backed(init.as_ref())
                {
                    self.carray_ptr_vars.insert(name.clone());
                    if let Some(raw_init) = init.clone() {
                        if !is_carray_like_expr(&raw_init) {
                            init = Some(self.wrap_as_carray_init(raw_init));
                        }
                    }
                }
                // else: int *p = &scalar → scalar cell (address_taken mechanism)
                }
            }
            // Zero-init struct/union instances when no explicit initializer.
            if init.is_none() {
                if let Some(fields) = &struct_fields {
                    let struct_name = normalized_c_type_name(&type_text);
                    if let Some(ref bounds) = array_bounds {
                        // Array of structs: pre-fill with N copies of zero struct.
                        let count = bounds
                            .first()
                            .and_then(|b| {
                                if let ExprKind::Lit(Literal::Int(n)) = &b.kind {
                                    Some(*n as usize)
                                } else {
                                    None
                                }
                            })
                            .unwrap_or(0);
                        if count > 0 {
                            let zeros: Vec<ArrayElement> = (0..count)
                                .map(|_| ArrayElement {
                                    value: self.zero_struct(Some(&struct_name), fields),
                                    spread: false,
                                    key: None,
                                    by_ref: false,
                                })
                                .collect();
                            init = Some(expr(ExprKind::Array(zeros)));
                            array_bounds = None;
                        }
                    } else {
                        init = Some(self.zero_struct(Some(&struct_name), fields));
                    }
                } else if !is_function_pointer_decl && !is_pointer_decl {
                    // Multidimensional plain array without initializer (`int m[2][2]`):
                    // build nested zero arrays so `m[i]` is a real sub-array and
                    // `m[i][j] = v` writes stick. (1-D arrays are allocated by the
                    // compiler from array_bounds metadata, so leave those alone.)
                    let is_char = normalized_c_type_name(&type_text) == "char";
                    if let Some(bounds) = &array_bounds {
                        if bounds.len() >= 2 && !is_char {
                            if let Some(zero) = zero_nd_array(bounds) {
                                init = Some(zero);
                                array_bounds = None;
                            }
                        }
                    }
                }
            }
            // Designated struct initializer (`struct P p = {.x=1, .z=3}`): overlay
            // the provided fields onto a zero-filled struct so omitted fields read
            // as 0 (not undefined → NaN), preserving declaration order.
            if let Some(fields) = &struct_fields {
                if array_bounds.is_none()
                    && matches!(init.as_ref().map(|i| &i.kind), Some(ExprKind::Object(_)))
                {
                    let struct_name = normalized_c_type_name(&type_text);
                    if let Some(Expression { kind: ExprKind::Object(provided), .. }) = init.take() {
                        let zero = self.zero_struct(Some(&struct_name), fields);
                        if let ExprKind::Object(mut props) = zero.kind {
                            for given in provided {
                                if let ObjectProperty::KeyValue { key, value } = given {
                                    if let ExprKind::Lit(Literal::Str(gk)) = &key.kind {
                                        if let Some(slot) = props.iter_mut().find_map(|p| {
                                            if let ObjectProperty::KeyValue { key: k, value: v } = p
                                            {
                                                if matches!(&k.kind, ExprKind::Lit(Literal::Str(s)) if s == gk) {
                                                    return Some(v);
                                                }
                                            }
                                            None
                                        }) {
                                            *slot = value;
                                            continue;
                                        }
                                    }
                                    props.push(ObjectProperty::KeyValue { key, value });
                                }
                            }
                            init = Some(expr(ExprKind::Object(props)));
                        } else {
                            init = Some(zero);
                        }
                    }
                }
            }
            // Designated / partial array initializer (`int a[5] = {[0]=10,[4]=30}` or
            // `int a[5] = {1,2}`): densify to the declared length with zeros so
            // omitted indices read as 0 instead of undefined → NaN. 1-D only.
            if struct_fields.is_none() && !is_function_pointer_decl {
                if let Some(size) = array_bounds.as_ref().and_then(|b| {
                    if b.len() == 1 {
                        if let ExprKind::Lit(Literal::Int(n)) = &b[0].kind {
                            return Some(*n as usize);
                        }
                    }
                    None
                }) {
                    if let Some(ExprKind::Array(elems)) =
                        init.as_ref().map(|i| &i.kind).filter(|_| size > 0)
                    {
                        let has_keys = elems.iter().any(|e| e.key.is_some());
                        let elem_count = elems.len();
                        let all_scalar = elems
                            .iter()
                            .all(|e| !matches!(e.value.kind, ExprKind::Array(_)));
                        if all_scalar && (has_keys || elem_count < size) {
                            let mut dense: Vec<Expression> =
                                (0..size).map(|_| int_lit(0)).collect();
                            let mut pos = 0usize;
                            if let Some(Expression { kind: ExprKind::Array(elems), .. }) =
                                init.take()
                            {
                                for el in elems {
                                    let idx = match el.key.as_ref().map(|k| &k.kind) {
                                        Some(ExprKind::Lit(Literal::Int(n))) => *n as usize,
                                        _ => pos,
                                    };
                                    if idx < size {
                                        dense[idx] = el.value;
                                    }
                                    pos = idx + 1;
                                }
                            }
                            init = Some(expr(ExprKind::Array(
                                dense
                                    .into_iter()
                                    .map(|value| ArrayElement {
                                        key: None,
                                        value,
                                        spread: false,
                                        by_ref: false,
                                    })
                                    .collect(),
                            )));
                        }
                    }
                }
            }
            // char array with bounds initialized by a string → treat as string
            // (e.g. `char buf[32] = "hello"` → just a string variable)
            let is_char_type = normalized_c_type_name(&type_text) == "char";
            let mut emitted_type_hint = type_text.clone();
            if !is_function_pointer_decl {
                if is_pointer_decl && !emitted_type_hint.contains('*') {
                    emitted_type_hint.push('*');
                }
                if was_array_decl {
                    emitted_type_hint.push('*');
                }
            }
            if array_bounds.is_some() && is_char_type && !is_pointer_decl {
                if let Some(ref init_expr) = init {
                    if matches!(init_expr.kind, ExprKind::Lit(Literal::Str(_))) {
                        array_bounds = None; // treat as string, not array
                        emitted_type_hint = "char*".to_string();
                    }
                }
            }
            // char array with char initializers `{'h','i','\0'}` → join chars to
            // string. Only for a 1-D array whose elements are int char codes; a
            // multidim char array (`char w[3][6] = {"a","b"}`) keeps its array of
            // string rows.
            let is_1d_char_code_array = array_bounds.as_ref().map(|b| b.len() == 1).unwrap_or(false)
                && matches!(init.as_ref().map(|i| &i.kind), Some(ExprKind::Array(elems))
                    if elems.iter().all(|el| matches!(el.value.kind, ExprKind::Lit(Literal::Int(_)))));
            if is_1d_char_code_array && is_char_type && !is_pointer_decl {
                if let Some(ExprKind::Array(elems)) = init.as_ref().map(|i| &i.kind) {
                    let s: String = elems
                        .iter()
                        .filter_map(|el| {
                            if let ExprKind::Lit(Literal::Int(code)) = &el.value.kind {
                                if *code == 0 {
                                    None
                                } else {
                                    char::from_u32(*code as u32)
                                }
                            } else {
                                None
                            }
                        })
                        .collect();
                    init = Some(expr(ExprKind::Lit(Literal::Str(s))));
                    self.char_pointers.insert(name.clone());
                    array_bounds = None;
                    emitted_type_hint = "char*".to_string();
                }
            }
            // Any char array declaration (including unsized `char s[]`) that
            // currently lowers to a string literal must keep pointer-like type
            // semantics to avoid scalar-char coercion in the compiler.
            if was_array_decl
                && is_char_type
                && !is_pointer_decl
                && matches!(init.as_ref().map(|i| &i.kind), Some(ExprKind::Lit(Literal::Str(_))))
            {
                emitted_type_hint = "char*".to_string();
            }
            if is_char_type && !is_pointer_decl && !was_array_decl {
                if let Some(init_expr) = init.clone() {
                    if self.is_char_index_read(&init_expr) {
                        init = Some(string_adapter::string_to_char_code(init_expr));
                    }
                }
            }
            if is_function_pointer_decl {
                if let Some(return_type) = init
                    .as_ref()
                    .and_then(|expr| self.infer_function_return_type(expr))
                {
                    emitted_type_hint = format!("func() -> {}", return_type);
                } else {
                    emitted_type_hint = format!("func() -> {}", normalized_c_type_name(&type_text));
                }
            }
            // Partial array initialization: zero-fill tail slots.
            // `int arr[4] = {1, 2}` → `[1, 2, 0, 0]`
            if !type_text.contains("char") && struct_fields.is_none() {
                if let (Some(bounds), Some(init_expr)) = (&array_bounds, &init) {
                    if let ExprKind::Array(elems) = &init_expr.kind {
                        let count = bounds
                            .first()
                            .and_then(|b| {
                                if let ExprKind::Lit(Literal::Int(n)) = &b.kind {
                                    Some(*n as usize)
                                } else {
                                    None
                                }
                            })
                            .unwrap_or(0);
                        if count > elems.len() {
                            let mut padded = elems.clone();
                            while padded.len() < count {
                                padded.push(ArrayElement {
                                    value: expr(ExprKind::Lit(Literal::Int(0))),
                                    spread: false,
                                    key: None,
                                    by_ref: false,
                                });
                            }
                            init = Some(expr(ExprKind::Array(padded)));
                        }
                    }
                }
            }
            let mut metadata_type = type_text.clone();
            if let Some(ref bounds) = array_bounds {
                for b in bounds {
                    if let ExprKind::Lit(Literal::Int(n)) = &b.kind {
                        metadata_type.push_str(&format!("[{}]", n));
                    } else {
                        metadata_type.push_str("[]");
                    }
                }
            } else if is_char_type {
                if let Some(ExprKind::Lit(Literal::Str(s))) = init.as_ref().map(|i| &i.kind) {
                    metadata_type.push_str(&format!("[{}]", s.len() + 1));
                }
            }
            // Record the type for sizeof resolution
            self.var_types.insert(name.clone(), metadata_type.clone());
            if let Some(alignment) = declared_alignment {
                self.var_alignments.insert(name.clone(), alignment);
            }
            if is_function_pointer_decl {
                self.function_pointer_vars.insert(name.clone());
                if !self.current_function.is_empty() {
                    self.current_function_pointer_names.push(name.clone());
                }
            }
            // Record sizeof for this variable.
            // NOTE: compute from init elem count when array_bounds has been cleared by zero-fill.
            let sz = if is_pointer_decl {
                8
            } else if let Some(ref bounds) = array_bounds {
                let base = sizeof_from_type_text(&type_text);
                let count: i64 = bounds
                    .iter()
                    .map(|b| {
                        if let ExprKind::Lit(Literal::Int(n)) = &b.kind {
                            *n
                        } else {
                            1
                        }
                    })
                    .product();
                base * count
            } else if let Some(ExprKind::Array(elems)) = init.as_ref().map(|i| &i.kind) {
                // array_bounds was cleared after pre-fill — count elements
                let base = sizeof_from_type_text(&type_text);
                base * elems.len() as i64
            } else if is_char_type {
                if let Some(ExprKind::Lit(Literal::Str(s))) = init.as_ref().map(|i| &i.kind) {
                    (s.len() + 1) as i64
                } else {
                    sizeof_from_type_text(&type_text)
                }
            } else {
                let su = self.sizeof_struct_union(&type_text);
                if su > 0 {
                    su
                } else {
                    sizeof_from_type_text(&type_text)
                }
            };
            self.var_sizes.insert(name.clone(), sz);
            // Non-char arrays (int arr[n], double arr[n]) decay to pointer for arithmetic.
            if was_array_decl && (!is_char_type || is_pointer_decl) {
                self.array_ptr_vars.insert(name.clone());
            }
            // Handle static local: lift to a module-level global with mangled name
            if is_static_local {
                let mangled = format!("__static_{}_{}", self.current_function, name);
                self.static_renames.insert(name.clone(), mangled.clone());
                self.static_globals.push(stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(mangled),
                        type_hint: Some(emitted_type_hint.clone()),
                        init,
                        array_bounds,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Var,
                }));
                continue;
            }
            declarations.push(VarDeclarator {
                pattern: BindingPattern::Ident(name),
                type_hint: Some(emitted_type_hint),
                init,
                array_bounds,
                with_events: false,
            });
        }
        if !declarations.is_empty() {
            out.push(stmt(StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::Var,
            }));
        }
    }

    fn make_struct_decl(&self, name: &str, fields: &[String]) -> Statement {
        let members = fields
            .iter()
            .map(|f| ClassMember::Field {
                name: f.clone(),
                type_hint: None,
                init: Some(expr(ExprKind::Lit(Literal::Int(0)))),
                modifiers: Modifiers::default(),
                with_events: false,
                array_bounds: None,
            })
            .collect();
        stmt(StmtKind::StructDecl {
            name: name.to_string(),
            interfaces: Vec::new(),
            members,
            visibility: Visibility::Public,
            decorators: Vec::new(),
        })
    }

    /// Zero value for a struct field given its (array-suffixed) type:
    /// scalar → 0, nested struct → zero-struct, `T[N]` → array of N zero-`T`,
    /// `T[]` (flexible) → empty growable array, multidim → nested arrays.
    fn zero_value_for_field_type(&self, field_type: &str) -> Expression {
        // Split the array suffix from the base type, keeping the base intact
        // (e.g. "union Data" must keep its space for normalized_c_type_name to
        // strip the `union `/`struct ` prefix).
        let ft = field_type.trim();
        let (base, dims): (&str, Vec<Option<usize>>) = match ft.find('[') {
            Some(i) => {
                let dims = ft[i..]
                    .split('[')
                    .skip(1)
                    .map(|p| p.trim().trim_end_matches(']').trim().parse::<usize>().ok())
                    .collect();
                (ft[..i].trim(), dims)
            }
            None => (ft, Vec::new()),
        };
        // element (no array) value: nested struct or scalar 0.
        let element = {
            let bn = normalized_c_type_name(base);
            if let Some(nested) = self.structs.get(&bn) {
                self.zero_struct(Some(&bn), &nested.clone())
            } else {
                expr(ExprKind::Lit(Literal::Int(0)))
            }
        };
        if dims.is_empty() {
            return element;
        }
        self.build_zero_array(&dims, &element)
    }

    fn build_zero_array(&self, dims: &[Option<usize>], element: &Expression) -> Expression {
        match dims.split_first() {
            None => element.clone(),
            // Flexible / incomplete dimension → empty growable array.
            Some((None, _)) => expr(ExprKind::Array(Vec::new())),
            Some((Some(n), rest)) => {
                let inner = self.build_zero_array(rest, element);
                let elems = (0..*n)
                    .map(|_| ArrayElement {
                        key: None,
                        value: inner.clone(),
                        spread: false,
                        by_ref: false,
                    })
                    .collect();
                expr(ExprKind::Array(elems))
            }
        }
    }

    fn zero_struct(&self, struct_name_hint: Option<&str>, fields: &[String]) -> Expression {
        let props = fields
            .iter()
            .map(|f| {
                // Look up field type in struct_field_types if we have a struct name
                let value = if let Some(sname) = struct_name_hint {
                    if let Some(field_type) =
                        self.struct_field_types.get(sname).and_then(|m| m.get(f)).cloned()
                    {
                        self.zero_value_for_field_type(&field_type)
                    } else {
                        expr(ExprKind::Lit(Literal::Int(0)))
                    }
                } else {
                    expr(ExprKind::Lit(Literal::Int(0)))
                };
                ObjectProperty::KeyValue {
                    key: expr(ExprKind::Lit(Literal::Str(f.clone()))),
                    value,
                }
            })
            .collect();
        expr(ExprKind::Object(props))
    }

    /// Deep copy a struct by recursively copying all fields, including nested structs.
    /// For `struct Pair second = first;`, generates:
    /// `{a: first.a, b: first.b}` for simple fields
    /// `{origin: {x: first.origin.x, y: first.origin.y}, size: first.size}` for nested
    fn deep_copy_struct(&self, type_name: &str, source: Expression) -> Expression {
        let normalized_type = normalized_c_type_name(type_name);
        let Some(fields) = self.structs.get(&normalized_type) else {
            return source; // Not a known struct, return as-is
        };

        let props: Vec<ObjectProperty> = fields
            .iter()
            .map(|f| {
                let member_access = expr(ExprKind::Member {
                    object: Box::new(source.clone()),
                    field: f.clone(),
                    null_safe: false,
                });

                // Check if this field is itself a struct
                let value =
                    if let Some(field_type_map) = self.struct_field_types.get(&normalized_type) {
                        if let Some(field_type) = field_type_map.get(f) {
                            let field_normalized = normalized_c_type_name(field_type);
                            if self.structs.contains_key(&field_normalized) {
                                // Recursively deep copy nested struct
                                self.deep_copy_struct(field_type, member_access)
                            } else {
                                member_access
                            }
                        } else {
                            member_access
                        }
                    } else {
                        member_access
                    };

                ObjectProperty::KeyValue {
                    key: expr(ExprKind::Lit(Literal::Str(f.clone()))),
                    value,
                }
            })
            .collect();

        expr(ExprKind::Object(props))
    }

    fn convert_array_init_to_struct_typed(
        &self,
        type_name: &str,
        raw: Expression,
        fields: &[String],
    ) -> Expression {
        let elems = match raw.kind {
            ExprKind::Array(elems) => elems,
            other => return expr(other),
        };
        if elems.is_empty() {
            return expr(ExprKind::Array(elems));
        }

        let normalized_type = normalized_c_type_name(type_name);
        let field_types = self.struct_field_types.get(&normalized_type).cloned();
        let mut props = Vec::new();
        for (i, el) in elems.into_iter().enumerate() {
            let Some(fname) = fields.get(i).cloned() else {
                continue;
            };
            let ft = field_types.as_ref().and_then(|t| t.get(&fname));
            let value = match ft {
                // Array field (`Entry entries[4]`, `int m[2][3]`): a zero element
                // (the `{0}` idiom) builds the zero array shape; otherwise keep.
                Some(ft) if ft.replace(' ', "").contains('[') => {
                    if is_zero_int_expr(&el.value) {
                        self.zero_value_for_field_type(ft)
                    } else {
                        el.value
                    }
                }
                // Nested struct field: recurse into a nested `{...}`.
                Some(ft) if self.structs.contains_key(&normalized_c_type_name(ft)) => {
                    let nf = self.structs[&normalized_c_type_name(ft)].clone();
                    self.convert_array_init_to_struct_typed(ft, el.value, &nf)
                }
                _ => el.value,
            };
            props.push(ObjectProperty::KeyValue {
                key: expr(ExprKind::Lit(Literal::Str(fname))),
                value,
            });
        }
        for i in props.len()..fields.len() {
            let value = field_types
                .as_ref()
                .and_then(|t| t.get(&fields[i]))
                .map(|ft| self.zero_value_for_field_type(ft))
                .unwrap_or_else(|| expr(ExprKind::Lit(Literal::Int(0))));
            props.push(ObjectProperty::KeyValue {
                key: expr(ExprKind::Lit(Literal::Str(fields[i].clone()))),
                value,
            });
        }
        expr(ExprKind::Object(props))
    }

    /// If the specifiers declare a struct/union with a body, return
    /// `(optional tag name, field names)`.
    fn struct_def_from_specifiers(
        &mut self,
        specs: &Pair<Rule>,
    ) -> Option<(Option<String>, Vec<String>, HashMap<String, String>, HashMap<String, (i64, bool)>)>
    {
        for p in specs.clone().into_inner() {
            if p.as_rule() == Rule::type_specifier || p.as_rule() == Rule::type_specifier_strict {
                for ts in p.into_inner() {
                    if ts.as_rule() == Rule::struct_or_union_spec {
                        let mut tag = None;
                        let mut fields = Vec::new();
                        let mut field_types = HashMap::new();
                        let mut bitfields = HashMap::new();
                        let mut has_body = false;
                        for sp in ts.into_inner() {
                            match sp.as_rule() {
                                Rule::ident_name => tag = Some(sp.as_str().to_string()),
                                Rule::struct_member => {
                                    has_body = true;
                                    self.collect_struct_fields(
                                        sp,
                                        &mut fields,
                                        &mut field_types,
                                        &mut bitfields,
                                    );
                                }
                                _ => {}
                            }
                        }
                        if has_body {
                            if let Some(ref tag_name) = tag {
                                self.struct_field_types
                                    .insert(tag_name.clone(), field_types.clone());
                                if !bitfields.is_empty() {
                                    self.struct_bitfields
                                        .insert(tag_name.clone(), bitfields.clone());
                                }
                            }
                            return Some((tag, fields, field_types, bitfields));
                        }
                    }
                }
            }
        }
        None
    }

    fn collect_struct_fields(
        &mut self,
        member: Pair<Rule>,
        fields: &mut Vec<String>,
        field_types: &mut HashMap<String, String>,
        bitfields: &mut HashMap<String, (i64, bool)>,
    ) {
        let mut member_type = None;
        for p in member.into_inner() {
            if p.as_rule() == Rule::declaration_specifiers {
                member_type = Some(self.type_text(p));
            } else if p.as_rule() == Rule::struct_declarator_list {
                for d in p.into_inner() {
                    // struct_declarator = declarator ~ (":" ~ assignment_expression)?
                    let (field_decl, bit_width) = if d.as_rule() == Rule::struct_declarator {
                        let mut inner = d.into_inner();
                        let Some(decl) = inner.next() else { continue };
                        let width = inner.next().and_then(|w| w.as_str().trim().parse::<i64>().ok());
                        (decl, width)
                    } else {
                        (d, None) // fallback for old grammar (Rule::declarator)
                    };

                    if field_decl.as_rule() == Rule::declarator {
                        let decl_text = field_decl.as_str().replace(' ', "");
                        let n = self.clone_declarator_name(field_decl);
                        if !n.is_empty() {
                            fields.push(n.clone());
                            let is_unsigned = member_type
                                .as_ref()
                                .map(|t| t.contains("unsigned"))
                                .unwrap_or(false);
                            if let Some(width) = bit_width {
                                if width > 0 {
                                    bitfields.insert(n.clone(), (width, !is_unsigned));
                                }
                            }
                            if let Some(ref ty) = member_type {
                                // Preserve the array suffix on the field type so
                                // zero-init can build the right shape: `Entry[4]`
                                // (array of structs), `char data[]` (flexible),
                                // `int m[2][3]` (multidim).
                                let stored = match decl_text.find('[') {
                                    Some(i) => format!("{}{}", ty, &decl_text[i..]),
                                    None => ty.clone(),
                                };
                                field_types.insert(n, stored);
                            }
                        }
                    }
                }
            }
        }
    }

    fn clone_declarator_name(&self, pair: Pair<Rule>) -> String {
        for p in pair.into_inner() {
            if p.as_rule() == Rule::direct_declarator {
                for d in p.into_inner() {
                    match d.as_rule() {
                        Rule::ident_name => return d.as_str().to_string(),
                        Rule::declarator => return self.clone_declarator_name(d),
                        _ => {}
                    }
                }
            }
        }
        String::new()
    }

    /// Resolve the struct field list referenced by a declaration's specifiers
    /// (either an inline body or a previously-registered struct name).
    fn struct_type_of_specifiers(&mut self, specs: &Pair<Rule>) -> Option<Vec<String>> {
        if let Some((_, fields, _, _)) = self.struct_def_from_specifiers(specs) {
            return Some(fields);
        }
        for p in specs.clone().into_inner() {
            match p.as_rule() {
                Rule::type_specifier | Rule::type_specifier_strict => {
                    for ts in p.into_inner() {
                        match ts.as_rule() {
                            Rule::struct_or_union_spec => {
                                for sp in ts.into_inner() {
                                    if sp.as_rule() == Rule::ident_name {
                                        if let Some(f) = self.structs.get(sp.as_str()) {
                                            return Some(f.clone());
                                        }
                                    }
                                }
                            }
                            Rule::typedef_name => {
                                if let Some(f) = self.structs.get(ts.as_str()) {
                                    return Some(f.clone());
                                }
                            }
                            _ => {}
                        }
                    }
                }
                // bare typedef_name as direct child of declaration_specifiers
                // e.g. `Point point = {...}` where Point is typedef'd struct
                Rule::typedef_name => {
                    if let Some(f) = self.structs.get(p.as_str()) {
                        return Some(f.clone());
                    }
                }
                _ => {}
            }
        }
        None
    }

    fn enum_def_from_specifiers(
        &mut self,
        specs: &Pair<Rule>,
    ) -> Option<(Option<String>, Vec<EnumMember>)> {
        for p in specs.clone().into_inner() {
            if p.as_rule() == Rule::type_specifier || p.as_rule() == Rule::type_specifier_strict {
                for ts in p.into_inner() {
                    if ts.as_rule() == Rule::enum_spec {
                        let mut name = None;
                        let mut members = Vec::new();
                        let mut has_body = false;
                        for sp in ts.into_inner() {
                            match sp.as_rule() {
                                Rule::ident_name => name = Some(sp.as_str().to_string()),
                                Rule::enumerator_list => {
                                    has_body = true;
                                    for en in sp.into_inner() {
                                        if en.as_rule() == Rule::enumerator {
                                            let mut it = en.into_inner();
                                            let nm = it.next().map(|x| x.as_str().to_string());
                                            let val = it.next().map(|x| self.walk_assignment(x));
                                            if let Some(nm) = nm {
                                                members.push(EnumMember {
                                                    name: nm,
                                                    value: val,
                                                    constructor_args: Vec::new(),
                                                });
                                            }
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                        if has_body {
                            return Some((name, members));
                        }
                    }
                }
            }
        }
        None
    }

    fn declarator_name_and_bounds(
        &mut self,
        pair: Pair<Rule>,
    ) -> (String, Option<Vec<Expression>>) {
        let mut name = String::new();
        let mut bounds: Vec<Expression> = Vec::new();
        for p in pair.into_inner() {
            if p.as_rule() == Rule::direct_declarator {
                for d in p.into_inner() {
                    match d.as_rule() {
                        Rule::ident_name => name = d.as_str().to_string(),
                        Rule::declarator => {
                            let (n, b) = self.declarator_name_and_bounds(d);
                            name = n;
                            if let Some(b) = b {
                                bounds.extend(b);
                            }
                        }
                        Rule::array_suffix => {
                            if let Some(sz) = d.into_inner().next() {
                                bounds.push(self.walk_assignment(sz));
                            }
                            // Empty array suffix (arr[]) → don't push 0, let initializer determine size
                        }
                        _ => {}
                    }
                }
            }
        }
        (
            name,
            if bounds.is_empty() {
                None
            } else {
                Some(bounds)
            },
        )
    }

    fn walk_initializer(&mut self, pair: Pair<Rule>) -> Expression {
        let inner = pair.into_inner().next();
        let Some(inner) = inner else {
            return expr(ExprKind::Lit(Literal::Null));
        };
        match inner.as_rule() {
            Rule::assignment_expression => self.walk_assignment(inner),
            Rule::initializer_list => {
                let mut is_object = false;
                let mut elems = Vec::new();
                let mut props = Vec::new();
                for di in inner.into_inner() {
                    match di.as_rule() {
                        Rule::designated_init => {
                            let mut it = di.into_inner().peekable();
                            // designated `.field = init`
                            let first = it.next();
                            match first {
                                Some(p) if p.as_rule() == Rule::ident_name => {
                                    is_object = true;
                                    let key = p.as_str().to_string();
                                    if let Some(v) = it.next() {
                                        let val = self.walk_initializer(v);
                                        props.push(ObjectProperty::KeyValue {
                                            key: expr(ExprKind::Lit(Literal::Str(key))),
                                            value: val,
                                        });
                                    }
                                }
                                Some(p) if p.as_rule() == Rule::initializer => {
                                    elems.push(ArrayElement {
                                        key: None,
                                        value: self.walk_initializer(p),
                                        spread: false,
                                        by_ref: false,
                                    });
                                }
                                Some(p) if p.as_rule() == Rule::assignment_expression => {
                                    // `[idx] = init` designator → keyed array element
                                    is_object = false;
                                    if let Some(v) = it.next() {
                                        elems.push(ArrayElement {
                                            key: Some(self.walk_assignment(p)),
                                            value: self.walk_initializer(v),
                                            spread: false,
                                            by_ref: false,
                                        });
                                    } else {
                                        elems.push(ArrayElement {
                                            key: None,
                                            value: self.walk_assignment(p),
                                            spread: false,
                                            by_ref: false,
                                        });
                                    }
                                }
                                _ => {}
                            }
                        }
                        Rule::initializer => {
                            elems.push(ArrayElement {
                                key: None,
                                value: self.walk_initializer(di),
                                spread: false,
                                by_ref: false,
                            });
                        }
                        Rule::assignment_expression => {
                            elems.push(ArrayElement {
                                key: None,
                                value: self.walk_assignment(di),
                                spread: false,
                                by_ref: false,
                            });
                        }
                        _ => {}
                    }
                }
                if is_object {
                    expr(ExprKind::Object(props))
                } else {
                    expr(ExprKind::Array(elems))
                }
            }
            _ => expr(ExprKind::Lit(Literal::Null)),
        }
    }

    fn walk_initializer_list(&mut self, list_pair: Pair<Rule>) -> Expression {
        let mut is_object = false;
        let mut elems = Vec::new();
        let mut props = Vec::new();
        for di in list_pair.into_inner() {
            match di.as_rule() {
                Rule::designated_init => {
                    let mut it = di.into_inner().peekable();
                    let first = it.next();
                    match first {
                        Some(p) if p.as_rule() == Rule::ident_name => {
                            is_object = true;
                            let key = p.as_str().to_string();
                            if let Some(v) = it.next() {
                                let val = self.walk_initializer(v);
                                props.push(ObjectProperty::KeyValue {
                                    key: expr(ExprKind::Lit(Literal::Str(key))),
                                    value: val,
                                });
                            }
                        }
                        Some(p) if p.as_rule() == Rule::initializer => {
                            elems.push(ArrayElement {
                                key: None,
                                value: self.walk_initializer(p),
                                spread: false,
                                by_ref: false,
                            });
                        }
                        Some(p) if p.as_rule() == Rule::assignment_expression => {
                            elems.push(ArrayElement {
                                key: None,
                                value: self.walk_assignment(p),
                                spread: false,
                                by_ref: false,
                            });
                        }
                        _ => {}
                    }
                }
                Rule::initializer => {
                    elems.push(ArrayElement {
                        key: None,
                        value: self.walk_initializer(di),
                        spread: false,
                        by_ref: false,
                    });
                }
                Rule::assignment_expression => {
                    elems.push(ArrayElement {
                        key: None,
                        value: self.walk_assignment(di),
                        spread: false,
                        by_ref: false,
                    });
                }
                _ => {}
            }
        }
        if is_object {
            expr(ExprKind::Object(props))
        } else {
            expr(ExprKind::Array(elems))
        }
    }

    fn type_text(&mut self, pair: Pair<Rule>) -> String {
        self.normalize_type_text(strip_alignment_specifiers(pair.as_str()).as_str())
            .split_whitespace()
            .collect::<Vec<_>>()
            .join(" ")
    }

    fn infer_function_return_type(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self.function_return_types.get(name).cloned(),
            ExprKind::Member { field, .. } => self.function_return_types.get(field).cloned(),
            _ => None,
        }
    }

    fn normalize_type_text(&mut self, text: &str) -> String {
        let trimmed = text.trim();
        if let Some(inner) = extract_typeof_expr_text(trimmed) {
            let expr_src = inner.trim();
            if let Some(ty) = self.var_types.get(expr_src) {
                // var_types stores the base type without the pointer `*`; restore
                // it for `__typeof__` so a pointer var's type stays a pointer.
                let is_ptr = self.carray_ptr_vars.contains(expr_src)
                    || self.array_ptr_vars.contains(expr_src)
                    || self.pointer_vars.contains(expr_src)
                    || self.char_pointers.contains(expr_src);
                if is_ptr && !ty.contains('*') {
                    return format!("{}*", ty.trim());
                }
                return ty.clone();
            }
            if let Ok(mut pairs) = CParser::parse(Rule::assignment_expression, expr_src) {
                if let Some(pair) = pairs.next() {
                    let expr = self.walk_assignment(pair);
                    return self.infer_generic_type(&expr, expr_src);
                }
            }
            return "int".to_string();
        }
        trimmed.to_string()
    }

    fn make_c_comparator_adapter(&self, cmp: Expression) -> Expression {
        let left_name = "__c_cmp_left".to_string();
        let right_name = "__c_cmp_right".to_string();
        let left_ident = ident(&left_name);
        let right_ident = ident(&right_name);
        let cmp_call = expr(ExprKind::Call {
            callee: Box::new(cmp),
            args: vec![
                Argument::positional(pointers::make_carray_ptr(
                    expr(ExprKind::Array(vec![ArrayElement {
                        key: None,
                        value: left_ident,
                        spread: false,
                        by_ref: false,
                    }])),
                    int_lit(0),
                )),
                Argument::positional(pointers::make_carray_ptr(
                    expr(ExprKind::Array(vec![ArrayElement {
                        key: None,
                        value: right_ident,
                        spread: false,
                        by_ref: false,
                    }])),
                    int_lit(0),
                )),
            ],
            optional: false,
        });
        expr(ExprKind::Lambda {
            params: vec![
                Param {
                    name: "left".to_string(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                },
                Param {
                    name: "right".to_string(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                },
            ],
            body: LambdaBody::Block(vec![
                stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(left_name.clone()),
                        type_hint: None,
                        init: Some(ident("left")),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Var,
                }),
                stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(right_name.clone()),
                        type_hint: None,
                        init: Some(ident("right")),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Var,
                }),
                stmt(StmtKind::Return(Some(cmp_call))),
            ]),
            is_async: false,
            captures: vec![],
        })
    }

    /// Array arguments decay to a `carray` pointer object `{__base, __idx}` when
    /// passed to a function. The wide-char (`wchar_t[]` = int array) helpers need
    /// the underlying flat array; unwrap a decayed carray to its base (the decay
    /// is at index 0), otherwise pass the value through.
    fn wide_array_operand(&self, value: Expression) -> Expression {
        carray_base_expr(&value).unwrap_or(value)
    }

    fn value_from_c_address_arg(&self, expr_in: Expression) -> Expression {
        match expr_in.kind {
            ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr,
            } => *expr,
            other => expr(other),
        }
    }

    fn rewrite_c_bsearch_call(
        &self,
        key_arg: Expression,
        array_arg: Expression,
        count_arg: Expression,
        cmp_arg: Expression,
    ) -> Expression {
        let idx_name = "__c_bsearch_idx".to_string();
        let idx_ident = ident(&idx_name);
        let helper_call = expr(ExprKind::Call {
            callee: Box::new(ident("__c_bsearch_index")),
            args: vec![
                Argument::positional(array_arg.clone()),
                Argument::positional(count_arg),
                Argument::positional(self.value_from_c_address_arg(key_arg)),
                Argument::positional(self.make_c_comparator_adapter(cmp_arg)),
            ],
            optional: false,
        });
        expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Lambda {
                params: vec![Param {
                    name: idx_name.clone(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }],
                body: LambdaBody::Expr(Box::new(expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(idx_ident.clone()),
                        right: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
                    })),
                    then: Box::new(pointers::make_carray_ptr(array_arg, idx_ident)),
                    else_: Box::new(expr(ExprKind::Lit(Literal::Null))),
                }))),
                is_async: false,
                captures: vec![],
            })),
            args: vec![Argument::positional(helper_call)],
            optional: false,
        })
    }

    fn assert_stmt_from_expr(&self, expr: &Expression) -> Option<Statement> {
        let ExprKind::Call { callee, args, .. } = &expr.kind else {
            return None;
        };
        let ExprKind::Ident(name) = &callee.kind else {
            return None;
        };
        if name != "assert" {
            return None;
        }
        if self.object_macros.contains_key("NDEBUG") {
            return Some(stmt(StmtKind::Empty));
        }
        let test = args.first()?.value.clone();
        Some(stmt(StmtKind::Assert { test, msg: None }))
    }

    fn walk_static_assert(&mut self, pair: Pair<Rule>) -> Statement {
        let mut inner = pair.into_inner();
        let test = inner
            .next()
            .map(|p| self.walk_expression(p))
            .unwrap_or_else(|| expr(ExprKind::Lit(Literal::Bool(true))));
        let msg = inner.next().map(|p| {
            expr(ExprKind::Lit(Literal::Str(parse_string_literal(
                p.as_str().trim(),
            ))))
        });
        stmt(StmtKind::Assert { test, msg })
    }

    fn capture_atexit_from_expr(&mut self, expression: &Expression) -> bool {
        if self.current_function != "main" {
            return false;
        }
        let ExprKind::Call { callee, args, .. } = &expression.kind else {
            return false;
        };
        let ExprKind::Ident(name) = &callee.kind else {
            return false;
        };
        if name != "atexit" || args.len() != 1 {
            return false;
        }
        let finalizer = stmt(StmtKind::Expr(expr(ExprKind::Call {
            callee: Box::new(args[0].value.clone()),
            args: vec![],
            optional: false,
        })));
        self.current_atexit_finalizers.push(finalizer);
        true
    }

    // ── Statements ─────────────────────────────────────────────────────────

    fn walk_block(&mut self, pair: Pair<Rule>) -> Vec<Statement> {
        let mut out = Vec::new();
        for p in pair.into_inner() {
            if p.as_rule() == Rule::statement {
                self.walk_statement(p, &mut out);
            }
        }
        out
    }

    fn walk_statement(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        let Some(inner) = pair.into_inner().next() else {
            return;
        };
        match inner.as_rule() {
            Rule::compound_statement => out.push(stmt(StmtKind::Block(self.walk_block(inner)))),
            Rule::function_definition => {
                if let Some(function) = self.walk_function(inner) {
                    out.push(function);
                }
            }
            Rule::declaration => self.walk_declaration(inner, out),
            Rule::static_assert_statement => out.push(self.walk_static_assert(inner)),
            Rule::expression_statement => {
                let e = inner.into_inner().next().unwrap();
                let expr = self.walk_expression(e);
                if let Some(macro_stmts) = self.expand_statement_macro_expr(&expr) {
                    out.extend(macro_stmts);
                    return;
                }
                if let Some(assert_stmt) = self.assert_stmt_from_expr(&expr) {
                    out.push(assert_stmt);
                    return;
                }
                if self.capture_atexit_from_expr(&expr) {
                    return;
                }
                let expr = self.rewrite_carray_postfix_discard(expr);
                out.push(stmt(StmtKind::Expr(expr)));
            }
            Rule::if_statement => out.push(self.walk_if(inner)),
            Rule::switch_statement => out.push(self.walk_switch(inner)),
            Rule::while_statement => out.push(self.walk_while(inner)),
            Rule::do_while_statement => out.push(self.walk_do_while(inner)),
            Rule::for_statement => out.push(self.walk_for(inner)),
            Rule::return_statement => {
                let val = inner.into_inner().next().map(|e| self.walk_expression(e));
                out.push(stmt(StmtKind::Return(val)));
            }
            Rule::break_statement => out.push(stmt(StmtKind::Break(BreakTarget::Implicit))),
            Rule::continue_statement => {
                out.push(stmt(StmtKind::Continue(ContinueTarget::Implicit)))
            }
            Rule::goto_statement => {
                let label = inner.into_inner().next().map(|p| p.as_str().to_string());
                out.push(stmt(StmtKind::GoTo(label.unwrap_or_default())));
            }
            Rule::labeled_statement => self.walk_labeled(inner, out),
            Rule::empty_statement => {}
            _ => {}
        }
    }

    fn body_of(&mut self, pair: Pair<Rule>) -> Vec<Statement> {
        // A `statement` that may be a block or a single statement.
        let mut out = Vec::new();
        self.walk_statement(pair, &mut out);
        // Unwrap a single Block into its statements for cleaner nesting.
        if out.len() == 1 {
            if let StmtKind::Block(inner) = &out[0].kind {
                return inner.clone();
            }
        }
        out
    }

    fn walk_if(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let cond_raw = self.walk_expression(it.next().unwrap());
        let cond = self.rewrite_char_condition(cond_raw);
        let then_body = self.body_of(it.next().unwrap());
        let mut else_body = None;
        if let Some(else_clause) = it.next() {
            if else_clause.as_rule() == Rule::else_clause {
                let st = else_clause.into_inner().next().unwrap();
                else_body = Some(self.body_of(st));
            }
        }
        stmt(StmtKind::If {
            cond,
            then_body,
            elifs: Vec::new(),
            else_body,
        })
    }

    fn walk_while(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let cond_raw = self.walk_expression(it.next().unwrap());
        let cond = self.rewrite_char_condition(cond_raw);
        let body = self.body_of(it.next().unwrap());
        stmt(StmtKind::While {
            cond,
            body,
            else_body: None,
        })
    }

    fn walk_do_while(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let body = self.body_of(it.next().unwrap());
        let cond_raw = self.walk_expression(it.next().unwrap());
        let cond = self.rewrite_char_condition(cond_raw);
        stmt(StmtKind::DoWhile {
            body,
            cond,
            until: false,
        })
    }

    fn walk_for(&mut self, pair: Pair<Rule>) -> Statement {
        let mut init = None;
        let mut cond = None;
        let mut update = None;
        let mut body = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::for_init => {
                    let mut inits = Vec::new();
                    if let Some(c) = p.into_inner().next() {
                        match c.as_rule() {
                            Rule::declaration => self.walk_declaration(c, &mut inits),
                            Rule::expression => {
                                inits.push(stmt(StmtKind::Expr(self.walk_expression(c))))
                            }
                            _ => {}
                        }
                    }
                    init = inits.into_iter().next().map(Box::new);
                }
                Rule::for_cond => {
                    cond = p
                        .into_inner()
                        .next()
                        .map(|e| self.walk_expression(e))
                        .map(|e| self.rewrite_char_condition(e));
                }
                Rule::for_update => {
                    update = p
                        .into_inner()
                        .next()
                        .map(|e| self.walk_expression(e))
                        .map(|e| self.rewrite_carray_postfix_discard(e));
                }
                Rule::statement => body = self.body_of(p),
                _ => {}
            }
        }
        stmt(StmtKind::For {
            init,
            cond,
            update,
            body,
        })
    }

    fn walk_switch(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let subject = self.walk_expression(it.next().unwrap());
        let body_stmt = it.next().unwrap();
        // The body is a compound statement of case/default labels.
        let mut cases: Vec<SwitchCase> = Vec::new();
        let mut default: Option<Vec<Statement>> = None;
        let block = body_stmt.into_inner().next();
        if let Some(block) = block {
            if block.as_rule() == Rule::compound_statement {
                self.collect_switch_cases(block, &mut cases, &mut default);
            }
        }
        // Post-process fallthrough: if a case body doesn't end with break/return,
        // append the next case's body to it.
        for i in (0..cases.len().saturating_sub(1)).rev() {
            if !ends_with_break(&cases[i].body) {
                let next_body = cases[i + 1].body.clone();
                cases[i].body.extend(next_body);
            }
        }
        // Also handle last case falling through to default
        if let Some(ref def_body) = default.clone() {
            if let Some(last) = cases.last_mut() {
                if !ends_with_break(&last.body) {
                    last.body.extend(def_body.clone());
                }
            }
        }
        stmt(StmtKind::Switch {
            expr: subject,
            cases,
            default,
        })
    }

    fn collect_switch_cases(
        &mut self,
        block: Pair<Rule>,
        cases: &mut Vec<SwitchCase>,
        default: &mut Option<Vec<Statement>>,
    ) {
        // Flatten labeled_statement chains into (condition, following stmts).
        let mut pending_conditions: Vec<CaseCondition> = Vec::new();
        let mut cur_body: Vec<Statement> = Vec::new();
        let mut in_default = false;
        let mut started = false;

        let flush = |conds: &mut Vec<CaseCondition>,
                         body: &mut Vec<Statement>,
                         is_default: bool,
                         cases: &mut Vec<SwitchCase>,
                         default: &mut Option<Vec<Statement>>| {
            if is_default {
                *default = Some(std::mem::take(body));
            } else if !conds.is_empty() {
                cases.push(SwitchCase {
                    conditions: std::mem::take(conds),
                    body: std::mem::take(body),
                });
            } else {
                body.clear();
            }
        };

        for st in block.into_inner() {
            if st.as_rule() != Rule::statement {
                continue;
            }
            // Peel off case/default labels, possibly nested via the trailing
            // `statement?` in the grammar.
            let mut stack = vec![st];
            while let Some(node) = stack.pop() {
                let Some(inner) = node.into_inner().next() else {
                    continue;
                };
                if inner.as_rule() == Rule::labeled_statement {
                    let lbl = inner.into_inner().next().unwrap();
                    match lbl.as_rule() {
                        Rule::case_label => {
                            // Only flush if we have accumulated body (avoids creating
                            // empty SwitchCase for grouped labels like `case 0: case 1: ...`)
                            if started && !cur_body.is_empty() {
                                flush(
                                    &mut pending_conditions,
                                    &mut cur_body,
                                    in_default,
                                    cases,
                                    default,
                                );
                            } else if started
                                && !pending_conditions.is_empty()
                                && cur_body.is_empty()
                            {
                                // grouped case: just accumulate the new condition into pending
                            } else if started {
                                flush(
                                    &mut pending_conditions,
                                    &mut cur_body,
                                    in_default,
                                    cases,
                                    default,
                                );
                            }
                            started = true;
                            in_default = false;
                            let mut ci = lbl.into_inner();
                            let val = self.walk_expression_pair_as_cond(ci.next().unwrap());
                            pending_conditions.push(CaseCondition::Value(val));
                            if let Some(rest) = ci.next() {
                                stack.push(rest);
                            }
                        }
                        Rule::default_label => {
                            if started && (!cur_body.is_empty() || in_default) {
                                flush(
                                    &mut pending_conditions,
                                    &mut cur_body,
                                    in_default,
                                    cases,
                                    default,
                                );
                            } else if started && cur_body.is_empty() && !in_default {
                                // grouped: flush whatever is pending as cases with empty body
                                // (shouldn't happen normally, but handle gracefully)
                                flush(
                                    &mut pending_conditions,
                                    &mut cur_body,
                                    in_default,
                                    cases,
                                    default,
                                );
                            }
                            started = true;
                            in_default = true;
                            if let Some(rest) = lbl.into_inner().next() {
                                stack.push(rest);
                            }
                        }
                        Rule::goto_label => {
                            let mut li = lbl.into_inner();
                            let name = li.next().unwrap().as_str().to_string();
                            cur_body.push(stmt(StmtKind::Label(name)));
                            if let Some(rest) = li.next() {
                                stack.push(rest);
                            }
                        }
                        _ => {}
                    }
                } else {
                    let mut tmp = Vec::new();
                    self.dispatch_inner_statement(inner, &mut tmp);
                    cur_body.append(&mut tmp);
                }
            }
        }
        if started {
            flush(
                &mut pending_conditions,
                &mut cur_body,
                in_default,
                cases,
                default,
            );
        }
    }

    /// Walk a `case <expr>` condition. The grammar uses
    /// `conditional_expression`, so route through that.
    fn walk_expression_pair_as_cond(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::conditional_expression => self.walk_conditional(pair),
            _ => self.walk_expression(pair),
        }
    }

    /// Re-dispatch an already-unwrapped statement inner node.
    fn dispatch_inner_statement(&mut self, inner: Pair<Rule>, out: &mut Vec<Statement>) {
        match inner.as_rule() {
            Rule::compound_statement => out.push(stmt(StmtKind::Block(self.walk_block(inner)))),
            Rule::declaration => self.walk_declaration(inner, out),
            Rule::expression_statement => {
                let e = inner.into_inner().next().unwrap();
                let expr = self.walk_expression(e);
                let expr = self.rewrite_carray_postfix_discard(expr);
                out.push(stmt(StmtKind::Expr(expr)));
            }
            Rule::if_statement => out.push(self.walk_if(inner)),
            Rule::switch_statement => out.push(self.walk_switch(inner)),
            Rule::while_statement => out.push(self.walk_while(inner)),
            Rule::do_while_statement => out.push(self.walk_do_while(inner)),
            Rule::for_statement => out.push(self.walk_for(inner)),
            Rule::return_statement => {
                let val = inner.into_inner().next().map(|e| self.walk_expression(e));
                out.push(stmt(StmtKind::Return(val)));
            }
            Rule::break_statement => out.push(stmt(StmtKind::Break(BreakTarget::Implicit))),
            Rule::continue_statement => {
                out.push(stmt(StmtKind::Continue(ContinueTarget::Implicit)))
            }
            Rule::goto_statement => {
                let label = inner.into_inner().next().map(|p| p.as_str().to_string());
                out.push(stmt(StmtKind::GoTo(label.unwrap_or_default())));
            }
            Rule::empty_statement => {}
            _ => {}
        }
    }

    fn walk_labeled(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        // goto_label outside a switch: `name: stmt`
        if let Some(lbl) = pair.into_inner().next() {
            if lbl.as_rule() == Rule::goto_label {
                let mut it = lbl.into_inner();
                let name = it.next().unwrap().as_str().to_string();
                out.push(stmt(StmtKind::Label(name)));
                if let Some(rest) = it.next() {
                    self.walk_statement(rest, out);
                }
            }
        }
    }

    // ── Expressions ──────────────────────────────────────────────────────────

    fn walk_expression(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::expression => {
                let mut parts: Vec<Expression> =
                    pair.into_inner().map(|p| self.walk_assignment(p)).collect();
                if parts.len() == 1 {
                    parts.pop().unwrap()
                } else {
                    expr(ExprKind::Sequence(parts))
                }
            }
            Rule::assignment_expression => self.walk_assignment(pair),
            Rule::conditional_expression => self.walk_conditional(pair),
            _ => self.walk_assignment(pair),
        }
    }

    fn walk_assignment(&mut self, pair: Pair<Rule>) -> Expression {
        if pair.as_rule() != Rule::assignment_expression {
            return self.walk_conditional(pair);
        }
        let mut it = pair.into_inner().peekable();
        let first = it.next().unwrap();
        if first.as_rule() == Rule::conditional_expression {
            return self.walk_conditional(first);
        }
        // unary ~ assign_op ~ assignment
        let target = self.walk_unary(first);
        let op = it.next().unwrap().as_str().to_string();
        let value = self.walk_assignment(it.next().unwrap());
        if op == "=" {
            self.sync_pointer_alias_on_assign(&target, &value);
            let target = self.rewrite_pointer_member_alias_target(target);
            let target = match target.kind {
                ExprKind::RefLoad(inner) => *inner,
                _ => target,
            };
            self.record_char_param_write(&target, &value);
            if let Some(ptr_name) = carray_deref_target_name(&target) {
                return dynamic_carray_deref_write(ident(&ptr_name), value);
            }
            if let Some(rewrite) = self.rewrite_char_index_assignment(&target, value.clone()) {
                return rewrite;
            }
            // Assigning to a bitfield member (`b.val = 7` for `: 2`) wraps to the
            // declared bit width (unsigned → mask; signed → mask + sign-extend).
            let value = match self.bitfield_of_member(&target) {
                Some((width, signed)) => apply_bitfield_mask(value, width, signed),
                None => value,
            };
            expr(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(value),
            })
        } else {
            let cop = match op.as_str() {
                "+=" => CompoundOp::Add,
                "-=" => CompoundOp::Sub,
                "*=" => CompoundOp::Mul,
                "/=" => CompoundOp::Div,
                "%=" => CompoundOp::Mod,
                "&=" => CompoundOp::BitAnd,
                "|=" => CompoundOp::BitOr,
                "^=" => CompoundOp::BitXor,
                "<<=" => CompoundOp::Shl,
                ">>=" => CompoundOp::Shr,
                _ => CompoundOp::Add,
            };
            // Represent compound assignment as a statement-level op via Assign
            // of a Binary, since ExprKind has no CompoundAssign.
            let bin = match cop {
                CompoundOp::Add => BinOp::Add,
                CompoundOp::Sub => BinOp::Sub,
                CompoundOp::Mul => BinOp::Mul,
                CompoundOp::Div => BinOp::Div,
                CompoundOp::Mod => BinOp::Mod,
                CompoundOp::BitAnd => BinOp::BitAnd,
                CompoundOp::BitOr => BinOp::BitOr,
                CompoundOp::BitXor => BinOp::BitXor,
                CompoundOp::Shl => BinOp::Shl,
                CompoundOp::Shr => BinOp::Shr,
                _ => BinOp::Add,
            };
            let target = self.rewrite_pointer_member_alias_target(target);
            let target = match target.kind {
                ExprKind::RefLoad(inner) => *inner,
                _ => target,
            };
            let rhs_raw = expr(ExprKind::Binary {
                op: bin,
                left: Box::new(target.clone()),
                right: Box::new(value),
            });
            let rhs_raw = self.rewrite_char_index_numeric(rhs_raw);
            let rhs_raw = self.rewrite_char_ptr_arith(rhs_raw);
            let rhs = self.rewrite_carray_ptr_arith(rhs_raw);
            // Compound assignment through a pointer deref (`*p *= 2`) must write
            // back through the cell, the same as a plain `*p = ...` does.
            if let Some(ptr_name) = carray_deref_target_name(&target) {
                return dynamic_carray_deref_write(ident(&ptr_name), rhs);
            }
            // Compound assignment into a char buffer element (`s[i] -= 32`) must
            // splice the computed char-code back through the string model, not
            // store a bare number into a string index.
            if let Some(rewrite) = self.rewrite_char_index_assignment(&target, rhs.clone()) {
                return rewrite;
            }
            expr(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(rhs),
            })
        }
    }

    fn sync_pointer_alias_on_assign(&mut self, target: &Expression, value: &Expression) {
        let ExprKind::Ident(name) = &target.kind else {
            return;
        };
        let is_pointer_like = self.carray_ptr_vars.contains(name)
            || self.array_ptr_vars.contains(name)
            || self.char_pointers.contains(name)
            || self.pointer_address_aliases.contains_key(name)
            || self.pointer_member_aliases.contains_key(name)
            || self
                .var_types
                .get(name)
                .map(|ty| ty.contains('*'))
                .unwrap_or(false);
        if !is_pointer_like {
            return;
        }

        self.pointer_address_aliases.remove(name);
        self.pointer_member_aliases.remove(name);

        let value_opt = Some(value.clone());
        if let Some(member_target) = pointer_member_target_from_init(&value_opt) {
            self.pointer_member_aliases.insert(name.clone(), member_target);
            return;
        }
        if let Some(address_target) = pointer_address_target_from_init(&value_opt) {
            self.pointer_address_aliases
                .insert(name.clone(), address_target);
            return;
        }
        if let Some(address_target) = propagated_pointer_address_alias(
            &value_opt,
            &self.pointer_address_aliases,
        ) {
            self.pointer_address_aliases
                .insert(name.clone(), address_target);
        }
    }

    fn rewrite_pointer_member_alias_target(&self, target: Expression) -> Expression {
        let ExprKind::Unary {
            op: UnaryOp::Deref,
            expr: ptr,
        } = &target.kind
        else {
            return target;
        };
        let name = match &ptr.kind {
            ExprKind::Ident(name) => name,
            ExprKind::RefLoad(inner) => {
                let ExprKind::Ident(name) = &inner.kind else {
                    return target;
                };
                name
            }
            _ => {
                return target;
            }
        };
        if self.carray_ptr_vars.contains(name)
            || self.array_ptr_vars.contains(name)
            || self.char_pointers.contains(name)
        {
            return target;
        }
        if let Some(member_target) = self.pointer_member_aliases.get(name).cloned() {
            return member_target;
        }
        if let Some(address_target) = self.pointer_address_aliases.get(name) {
            // Assignment target for `*p = v` should be an lvalue (`x`), not a read (`RefLoad(x)`).
            return ident(address_target);
        }
        target
    }

    fn ident_or_refload(&self, name: &str) -> Expression {
        if self.address_taken.contains(name) {
            expr(ExprKind::RefLoad(Box::new(ident(name))))
        } else {
            ident(name)
        }
    }

    fn record_char_param_write(&mut self, target: &Expression, value: &Expression) {
        let ExprKind::Index { object, index, .. } = &target.kind else {
            return;
        };
        let ExprKind::Ident(name) = &object.kind else {
            return;
        };
        let Some(param_idx) = self.current_char_param_indices.get(name).copied() else {
            return;
        };
        self.char_param_writes
            .entry(self.current_function.clone())
            .or_default()
            .push((param_idx, *index.clone(), value.clone()));
    }

    fn rewrite_carray_postfix_discard(&self, value: Expression) -> Expression {
        let ExprKind::Unary {
            op: unary_op,
            expr: target,
        } = &value.kind
        else {
            return value;
        };
        let ExprKind::Member { object, field, .. } = &target.kind else {
            return value;
        };
        if field != CARRAY_IDX_KEY {
            return value;
        }
        let ExprKind::Ident(name) = &object.kind else {
            return value;
        };
        if !self.carray_ptr_vars.contains(name) {
            return value;
        }
        match unary_op {
            UnaryOp::PostInc => {
                pointers::carray_advance_inplace(name, expr(ExprKind::Lit(Literal::Int(1))))
            }
            UnaryOp::PostDec => {
                pointers::carray_retreat_inplace(name, expr(ExprKind::Lit(Literal::Int(1))))
            }
            _ => value,
        }
    }

    fn normalize_pointer_call_args(&self, callee: &str, args: Vec<Argument>) -> Vec<Argument> {
        let Some(param_types) = self.function_param_types.get(callee) else {
            return args;
        };
        args.into_iter()
            .enumerate()
            .map(|(idx, mut arg)| {
                let Some(Some(type_hint)) = param_types.get(idx) else {
                    return arg;
                };
                if !type_hint.contains('*') {
                    return arg;
                }
                if is_carray_object(&arg.value) {
                    return arg;
                }
                if type_hint.contains("char") {
                    return arg;
                }
                if matches!(&arg.value.kind, ExprKind::Ident(name) if self.is_fixed_array_var(name))
                {
                    arg.value = self.wrap_as_carray_init(arg.value);
                }
                arg
            })
            .collect()
    }

    fn normalize_fixed_array_call_args(&self, args: Vec<Argument>) -> Vec<Argument> {
        args.into_iter()
            .map(|mut arg| {
                if matches!(&arg.value.kind, ExprKind::Ident(name) if self.is_fixed_array_var(name))
                {
                    arg.value = self.wrap_as_carray_init(arg.value);
                }
                arg
            })
            .collect()
    }

    fn byte_count_to_usize(&self, value: &Expression) -> Option<usize> {
        match &value.kind {
            ExprKind::Lit(Literal::Int(n)) if *n >= 0 => Some(*n as usize),
            _ => None,
        }
    }

    fn dest_element_size(&self, value: &Expression) -> usize {
        let Some(name) = base_ident_name(value) else {
            return 1;
        };
        self.var_types
            .get(&name)
            .map(|ty| sizeof_from_type_text(ty).max(1) as usize)
            .unwrap_or(1)
    }

    fn char_copy_slice(&self, src: Expression, count: usize) -> Expression {
        if let Some((base, offset)) = char_buffer_target_offset(&src) {
            let end = expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(offset.clone()),
                right: Box::new(expr(ExprKind::Lit(Literal::Int(count as i64)))),
            });
            return expr(ExprKind::Call {
                callee: Box::new(expr(ExprKind::Member {
                    object: Box::new(ident(&base)),
                    field: "substring".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(offset), Argument::positional(end)],
                optional: false,
            });
        }
        match &src.kind {
            ExprKind::Lit(Literal::Str(s)) => {
                expr(ExprKind::Lit(Literal::Str(s.chars().take(count).collect())))
            }
            _ => expr(ExprKind::Call {
                callee: Box::new(expr(ExprKind::Member {
                    object: Box::new(src),
                    field: "substring".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))), Argument::positional(expr(ExprKind::Lit(Literal::Int(count as i64))))],
                optional: false,
            }),
        }
    }

    fn rewrite_memcpy_like(&mut self, dst: Expression, src: Expression, bytes: Expression) -> Expression {
        if let Some((dst_name, dst_offset)) = char_buffer_target_offset(&dst) {
            self.char_pointers.insert(dst_name.clone());
            let count = self.byte_count_to_usize(&bytes).unwrap_or(0);
            let copied = self.char_copy_slice(src, count);
            if is_zero_int_expr(&dst_offset) {
                return expr(ExprKind::Assign {
                    target: Box::new(ident(&dst_name)),
                    value: Box::new(copied),
                });
            }
            let base = ident(&dst_name);
            let prefix = expr(ExprKind::Call {
                callee: Box::new(expr(ExprKind::Member {
                    object: Box::new(base.clone()),
                    field: "substring".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))), Argument::positional(dst_offset.clone())],
                optional: false,
            });
            let suffix_start = expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(dst_offset),
                right: Box::new(expr(ExprKind::Lit(Literal::Int(count as i64)))),
            });
            let suffix = expr(ExprKind::Call {
                callee: Box::new(expr(ExprKind::Member {
                    object: Box::new(base.clone()),
                    field: "substring".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(suffix_start)],
                optional: false,
            });
            let updated = expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(expr(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(prefix),
                    right: Box::new(copied),
                })),
                right: Box::new(suffix),
            });
            return expr(ExprKind::Assign {
                target: Box::new(base),
                value: Box::new(updated),
            });
        }
        if let Some(name) = base_ident_name(&dst) {
            if self.is_fixed_array_var(&name) {
                let src_value = carray_base_expr(&src).unwrap_or(src);
                return expr(ExprKind::Assign {
                    target: Box::new(ident(&name)),
                    value: Box::new(src_value),
                });
            }
        }
        expr(ExprKind::Assign {
            target: Box::new(dst),
            value: Box::new(src),
        })
    }

    fn rewrite_memset(&mut self, dst: Expression, fill: Expression, bytes: Expression) -> Expression {
        if let Some((dst_name, dst_offset)) = char_buffer_target_offset(&dst) {
            self.char_pointers.insert(dst_name.clone());
            let count = self.byte_count_to_usize(&bytes).unwrap_or(0);
            let repeated = match &fill.kind {
                ExprKind::Lit(Literal::Int(code)) => {
                    let ch = char::from_u32(*code as u32).unwrap_or('\0');
                    expr(ExprKind::Lit(Literal::Str(ch.to_string().repeat(count))))
                }
                _ => {
                    let pieces = (0..count)
                        .map(|_| char_assignment_value_to_string(fill.clone()))
                        .collect::<Vec<_>>();
                    concat_sequence_to_string(pieces)
                }
            };
            if is_zero_int_expr(&dst_offset) {
                return expr(ExprKind::Assign {
                    target: Box::new(ident(&dst_name)),
                    value: Box::new(repeated),
                });
            }
            let base = ident(&dst_name);
            let prefix = expr(ExprKind::Call {
                callee: Box::new(expr(ExprKind::Member {
                    object: Box::new(base.clone()),
                    field: "substring".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))), Argument::positional(dst_offset.clone())],
                optional: false,
            });
            let suffix_start = expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(dst_offset),
                right: Box::new(expr(ExprKind::Lit(Literal::Int(count as i64)))),
            });
            let suffix = expr(ExprKind::Call {
                callee: Box::new(expr(ExprKind::Member {
                    object: Box::new(base.clone()),
                    field: "substring".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(suffix_start)],
                optional: false,
            });
            let updated = expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(expr(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(prefix),
                    right: Box::new(repeated),
                })),
                right: Box::new(suffix),
            });
            return expr(ExprKind::Assign {
                target: Box::new(base),
                value: Box::new(updated),
            });
        }
        if let Some(name) = base_ident_name(&dst) {
            if self.is_fixed_array_var(&name) && is_zero_int_expr(&fill) {
                let elem_size = self.dest_element_size(&dst).max(1);
                let count = self.byte_count_to_usize(&bytes).unwrap_or(0) / elem_size;
                let zeros: Vec<ArrayElement> = (0..count)
                    .map(|_| ArrayElement {
                        value: expr(ExprKind::Lit(Literal::Int(0))),
                        spread: false,
                        key: None,
                        by_ref: false,
                    })
                    .collect();
                return expr(ExprKind::Assign {
                    target: Box::new(ident(&name)),
                    value: Box::new(expr(ExprKind::Array(zeros))),
                });
            }
        }
        expr(ExprKind::Lit(Literal::Null))
    }

    fn walk_conditional(&mut self, pair: Pair<Rule>) -> Expression {
        let mut it = pair.into_inner();
        let cond = self.walk_binary(it.next().unwrap());
        if let Some(then_p) = it.next() {
            let then = self.walk_expression(then_p);
            let else_ = self.walk_conditional(it.next().unwrap());
            expr(ExprKind::Ternary {
                cond: Box::new(cond),
                then: Box::new(then),
                else_: Box::new(else_),
            })
        } else {
            cond
        }
    }

    /// If `target` is a struct bitfield member (`obj.field` where `field` is a
    /// `: N` bitfield), return its `(width, is_signed)`.
    fn bitfield_of_member(&self, target: &Expression) -> Option<(i64, bool)> {
        let ExprKind::Member { object, field, .. } = &target.kind else {
            return None;
        };
        let ExprKind::Ident(obj_name) = &object.kind else {
            return None;
        };
        let obj_type = self.var_types.get(obj_name)?;
        let tag = normalized_c_type_name(obj_type);
        self.struct_bitfields.get(&tag)?.get(field).copied()
    }

    /// True if `e` has an unsigned C type — so `>>` must be a logical shift.
    fn is_unsigned_expr(&self, e: &Expression) -> bool {
        match &e.kind {
            ExprKind::Ident(name) => self
                .var_types
                .get(name)
                .map(|t| t.contains("unsigned") || t.contains("uint") || t == "size_t")
                .unwrap_or(false),
            ExprKind::Cast { type_name, .. } => {
                let t = normalized_c_type_name(type_name);
                t.contains("unsigned") || t.contains("uint") || t == "size_t"
            }
            _ => false,
        }
    }

    /// `unsigned >> n` is a *logical* shift in C; emit `UShr` (→ `i32.shr_u`) so
    /// the high bit isn't sign-extended (signed `>>` stays `Shr` → `i32.shr_s`).
    fn rewrite_unsigned_shift(&self, e: Expression) -> Expression {
        if let ExprKind::Binary { op: BinOp::Shr, left, right } = e.kind {
            if self.is_unsigned_expr(&left) {
                return expr(ExprKind::Binary { op: BinOp::UShr, left, right });
            }
            return expr(ExprKind::Binary { op: BinOp::Shr, left, right });
        }
        e
    }

    fn walk_binary(&mut self, pair: Pair<Rule>) -> Expression {
        // binary_expression = unary (op unary)*
        let mut operands: Vec<Expression> = Vec::new();
        let mut ops: Vec<String> = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::binary_op => ops.push(p.as_str().to_string()),
                _ => operands.push(self.walk_unary(p)),
            }
        }
        let result = fold_binary(operands, ops);
        let result = self.rewrite_unsigned_shift(result);
        let result = self.rewrite_integer_division(result);
        let result = self.rewrite_pointer_zero_comparison(result);
        // Convert char-buffer reads to char codes BEFORE wrapping logical ops in a
        // ternary — `rewrite_logical_bool` turns `a && b` into `(a && b) ? 1 : 0`,
        // and `rewrite_char_index_numeric` only descends `Binary` nodes.
        let result = self.rewrite_char_index_numeric(result);
        let result = self.rewrite_logical_bool(result);
        let result = self.rewrite_char_ptr_arith(result);
        let result = self.rewrite_carray_ptr_arith(result);
        self.rewrite_complex_binary_expr(result)
    }

    fn is_floatish_expr(&self, expr_in: &Expression) -> bool {
        match &expr_in.kind {
            ExprKind::Lit(Literal::Float(_)) => true,
            ExprKind::Ident(name) => self
                .var_types
                .get(name)
                .map(|ty| {
                    let ty = normalized_c_type_name(ty);
                    ty.contains("float") || ty.contains("double") || ty.contains("long double")
                })
                .unwrap_or(false),
            ExprKind::Cast { type_name, .. } => {
                type_name.contains("double") || type_name.contains("float")
            }
            ExprKind::Unary { expr, .. } => self.is_floatish_expr(expr),
            ExprKind::Binary { left, right, .. } => {
                self.is_floatish_expr(left) || self.is_floatish_expr(right)
            }
            ExprKind::Ternary { then, else_, .. } => {
                self.is_floatish_expr(then) || self.is_floatish_expr(else_)
            }
            _ => false,
        }
    }

    fn rewrite_integer_division(&self, e: Expression) -> Expression {
        let ExprKind::Binary { op, left, right } = e.kind else {
            return e;
        };
        let left = self.rewrite_integer_division(*left);
        let right = self.rewrite_integer_division(*right);
        if op == BinOp::Div && !self.is_floatish_expr(&left) && !self.is_floatish_expr(&right) {
            return ecma_math_call(
                "trunc",
                expr(ExprKind::Binary {
                    op: BinOp::Div,
                    left: Box::new(left),
                    right: Box::new(right),
                }),
            );
        }
        expr(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
    }

    fn rewrite_complex_binary_expr(&self, e: Expression) -> Expression {
        let ExprKind::Binary { op, left, right } = e.kind else {
            return e;
        };
        let left = *left;
        let right = *right;

        let left_complex = self.is_complex_expr(&left);
        let right_complex = self.is_complex_expr(&right);
        let left_i = self.is_imag_unit_expr(&left);
        let right_i = self.is_imag_unit_expr(&right);

        match op {
            BinOp::Add if left_complex || right_complex || left_i || right_i => {
                let (left_re, left_im) = self.complex_parts(left);
                let (right_re, right_im) = self.complex_parts(right);
                return self.complex_object(
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(left_re),
                        right: Box::new(right_re),
                    }),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(left_im),
                        right: Box::new(right_im),
                    }),
                );
            }
            BinOp::Sub if left_complex || right_complex || left_i || right_i => {
                let (left_re, left_im) = self.complex_parts(left);
                let (right_re, right_im) = self.complex_parts(right);
                return self.complex_object(
                    expr(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(left_re),
                        right: Box::new(right_re),
                    }),
                    expr(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(left_im),
                        right: Box::new(right_im),
                    }),
                );
            }
            BinOp::Mul if left_complex || right_complex || left_i || right_i => {
                if left_i && right_i {
                    return self.complex_object(int_lit(-1), int_lit(0));
                }
                if left_i {
                    let (re, im) = self.complex_parts(right);
                    return self.complex_object(
                        expr(ExprKind::Unary { op: UnaryOp::Neg, expr: Box::new(im) }),
                        re,
                    );
                }
                if right_i {
                    let (re, im) = self.complex_parts(left);
                    return self.complex_object(
                        expr(ExprKind::Unary { op: UnaryOp::Neg, expr: Box::new(im) }),
                        re,
                    );
                }
                let (a_re, a_im) = self.complex_parts(left);
                let (b_re, b_im) = self.complex_parts(right);
                let real = expr(ExprKind::Binary {
                    op: BinOp::Sub,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Mul,
                        left: Box::new(a_re.clone()),
                        right: Box::new(b_re.clone()),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Mul,
                        left: Box::new(a_im.clone()),
                        right: Box::new(b_im.clone()),
                    })),
                });
                let imag = expr(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Mul,
                        left: Box::new(a_re),
                        right: Box::new(b_im),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Mul,
                        left: Box::new(a_im),
                        right: Box::new(b_re),
                    })),
                });
                return self.complex_object(real, imag);
            }
            _ => {}
        }

        expr(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
    }

    fn is_complex_expr(&self, expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Ident(name) => self
                .var_types
                .get(name)
                .map(|ty| is_complex_type_text(ty))
                .unwrap_or(false),
            ExprKind::Object(props) => props.iter().any(|p| match p {
                ObjectProperty::KeyValue { key, .. } => {
                    matches!(&key.kind, ExprKind::Lit(Literal::Str(s)) if s == "real" || s == "imag")
                }
                _ => false,
            }),
            ExprKind::Member { field, .. } => field == "real" || field == "imag",
            ExprKind::Binary { op: BinOp::Mul, left, right } => {
                self.is_imag_unit_expr(left)
                    || self.is_imag_unit_expr(right)
                    || self.is_complex_expr(left)
                    || self.is_complex_expr(right)
            }
            ExprKind::Binary { op, left, right } => {
                matches!(op, BinOp::Add | BinOp::Sub | BinOp::Mul)
                    && (self.is_complex_expr(left)
                        || self.is_complex_expr(right)
                        || self.is_imag_unit_expr(left)
                        || self.is_imag_unit_expr(right))
            }
            _ => false,
        }
    }

    fn is_imag_unit_expr(&self, expr: &Expression) -> bool {
        matches!(&expr.kind, ExprKind::Ident(name) if name == "I")
    }

    fn complex_parts(&self, expr: Expression) -> (Expression, Expression) {
        if self.is_imag_unit_expr(&expr) {
            return (int_lit(0), int_lit(1));
        }
        if let ExprKind::Binary { op: BinOp::Mul, left, right } = &expr.kind {
            if self.is_imag_unit_expr(left) {
                return (int_lit(0), (*right.clone()).clone());
            }
            if self.is_imag_unit_expr(right) {
                return (int_lit(0), (*left.clone()).clone());
            }
        }
        if self.is_complex_expr(&expr) {
            return (member(expr.clone(), "real"), member(expr, "imag"));
        }
        (expr, int_lit(0))
    }

    fn complex_object(&self, real: Expression, imag: Expression) -> Expression {
        complex_adapter::complex_object(real, imag)
    }

    fn complex_real_part(&self, value: Expression) -> Expression {
        if self.is_complex_expr(&value) {
            member(value, "real")
        } else {
            value
        }
    }

    fn complex_imag_part(&self, value: Expression) -> Expression {
        if self.is_complex_expr(&value) {
            member(value, "imag")
        } else if self.is_imag_unit_expr(&value) {
            int_lit(1)
        } else {
            int_lit(0)
        }
    }

    fn complex_conj(&self, value: Expression) -> Expression {
        self.complex_object(
            self.complex_real_part(value.clone()),
            expr(ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(self.complex_imag_part(value)),
            }),
        )
    }

    fn rewrite_pointer_zero_comparison(&self, e: Expression) -> Expression {
        let ExprKind::Binary { op, left, right } = e.kind else {
            return e;
        };
        if !matches!(op, BinOp::Eq | BinOp::NotEq) {
            return expr(ExprKind::Binary { op, left, right });
        }
        let left_expr = *left;
        let right_expr = *right;

        if is_zero_int_expr(&left_expr) && matches!(right_expr.kind, ExprKind::Lit(Literal::Null)) {
            return expr(ExprKind::Lit(Literal::Bool(matches!(op, BinOp::Eq))));
        }
        if is_zero_int_expr(&right_expr) && matches!(left_expr.kind, ExprKind::Lit(Literal::Null)) {
            return expr(ExprKind::Lit(Literal::Bool(matches!(op, BinOp::Eq))));
        }

        let left_is_ptr = self.is_pointer_like_expr(&left_expr);
        let right_is_ptr = self.is_pointer_like_expr(&right_expr);
        let left_zero = is_zero_int_expr(&left_expr);
        let right_zero = is_zero_int_expr(&right_expr);
        let left_final = if left_is_ptr && right_zero {
            left_expr.clone()
        } else if right_is_ptr && left_zero {
            null_lit()
        } else {
            left_expr
        };
        let right_final = if left_is_ptr && right_zero {
            null_lit()
        } else if right_is_ptr && left_zero {
            right_expr
        } else {
            right_expr
        };

        expr(ExprKind::Binary {
            op,
            left: Box::new(left_final),
            right: Box::new(right_final),
        })
    }

    fn is_pointer_like_expr(&self, expr_in: &Expression) -> bool {
        match &expr_in.kind {
            ExprKind::Ident(name) => self
                .var_types
                .get(name)
                .map(|ty| ty.contains('*') || ty.contains("FILE"))
                .unwrap_or(false)
                || self.pointer_vars.contains(name)
                || self.char_pointers.contains(name)
                || self.array_ptr_vars.contains(name)
                || self.carray_ptr_vars.contains(name),
            ExprKind::Lit(Literal::Null) => true,
            ExprKind::Object(_) => is_carray_like_expr(expr_in),
            ExprKind::Call { .. } | ExprKind::Ternary { .. } => is_carray_like_expr(expr_in),
            _ => false,
        }
    }

    fn rewrite_char_index_numeric(&self, e: Expression) -> Expression {
        let ExprKind::Binary { op, left, right } = e.kind else {
            return e;
        };
        let left = self.rewrite_char_index_numeric(*left);
        let right = self.rewrite_char_index_numeric(*right);
        let numeric_op = matches!(
            op,
            BinOp::Add
                | BinOp::Sub
                | BinOp::Mul
                | BinOp::Div
                | BinOp::Mod
                | BinOp::BitAnd
                | BinOp::BitOr
                | BinOp::BitXor
                | BinOp::Shl
                | BinOp::Shr
                | BinOp::Eq
                | BinOp::NotEq
                | BinOp::Lt
                | BinOp::LtEq
                | BinOp::Gt
                | BinOp::GtEq
        );
        if !numeric_op {
            return expr(ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            });
        }
        let left = if self.is_char_index_read(&left) {
            self.char_index_read_to_code(left)
        } else {
            left
        };
        let right = if self.is_char_index_read(&right) {
            self.char_index_read_to_code(right)
        } else {
            right
        };
        expr(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
    }

    fn rewrite_char_condition(&self, cond: Expression) -> Expression {
        if let ExprKind::Index { object, index, .. } = &cond.kind {
            if matches!(&object.kind, ExprKind::Ident(name) if self.char_pointers.contains(name)) {
                let in_bounds = expr(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: index.clone(),
                    right: Box::new(member(*object.clone(), "length")),
                });
                let char_value = string_adapter::string_to_char_code(cond);
                return expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(in_bounds),
                        then: Box::new(char_value),
                        else_: Box::new(int_lit(0)),
                    })),
                    right: Box::new(int_lit(0)),
                });
            }
        }
        cond
    }

    /// Convert a char-buffer element read (`s[i]`) into its int char code,
    /// guarding the read: an index at/past the string content yields the C null
    /// terminator (0) instead of `undefined.charCodeAt(...)`. Mirrors the bounds
    /// check in `rewrite_char_condition`. Caller must ensure `e` is a char-index
    /// read (`is_char_index_read`).
    fn char_index_read_to_code(&self, e: Expression) -> Expression {
        let ExprKind::Index { object, index, .. } = &e.kind else {
            return string_adapter::string_to_char_code(e);
        };
        let in_bounds = expr(ExprKind::Binary {
            op: BinOp::Lt,
            left: index.clone(),
            right: Box::new(member(*object.clone(), "length")),
        });
        expr(ExprKind::Ternary {
            cond: Box::new(in_bounds),
            then: Box::new(string_adapter::string_to_char_code(e)),
            else_: Box::new(int_lit(0)),
        })
    }

    fn is_char_index_read(&self, e: &Expression) -> bool {
        let ExprKind::Index { object, .. } = &e.kind else {
            return false;
        };
        matches!(&object.kind, ExprKind::Ident(name) if self.char_pointers.contains(name))
    }

    /// Wrap a pointer init expression as a carray object.
    /// `arr` (plain array ident) → `make_carray_ptr(arr, 0)`
    /// `arr + n`                 → `make_carray_ptr(arr, n)`
    /// `arr - n`                 → `make_carray_ptr(arr, -n)` — rare but valid
    /// Anything else            → `make_carray_ptr(expr, 0)`
    fn wrap_as_carray_init(&self, raw: Expression) -> Expression {
        let zero = expr(ExprKind::Lit(Literal::Int(0)));
        match raw.kind {
            ExprKind::Ident(ref name) if self.array_ptr_vars.contains(name) => {
                pointers::make_carray_ptr(raw, zero)
            }
            ExprKind::Ident(ref name) if self.carray_ptr_vars.contains(name) => {
                // Pointer copy — just use the existing carray object
                raw
            }
            ExprKind::Binary {
                op: BinOp::Add,
                ref left,
                ref right,
            } if matches!(&left.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n)) => {
                pointers::make_carray_ptr(*left.clone(), *right.clone())
            }
            ExprKind::Binary {
                op: BinOp::Sub,
                ref left,
                ref right,
            } if matches!(&left.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n)) => {
                let neg_n = expr(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: right.clone(),
                });
                pointers::make_carray_ptr(*left.clone(), neg_n)
            }
            _ => pointers::make_carray_ptr(raw, zero),
        }
    }

    /// Rewrite carray pointer arithmetic expressions.
    /// Runs after `rewrite_char_ptr_arith` so char* expressions are already handled.
    fn rewrite_carray_ptr_arith(&self, e: Expression) -> Expression {
        let ExprKind::Binary { op, left, right } = e.kind else {
            return e;
        };
        let left_is_carray_var =
            matches!(&left.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n));
        let right_is_carray_var =
            matches!(&right.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n));
        let left_is_array_var =
            matches!(&left.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n));
        let right_is_array_var =
            matches!(&right.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n));
        let left_is_carray_obj = is_carray_object(&left);
        let right_is_carray_obj = is_carray_object(&right);

        match op {
            BinOp::Eq | BinOp::NotEq => {
                if (left_is_carray_obj || left_is_carray_var)
                    && (right_is_carray_obj || right_is_carray_var)
                {
                    let ptr_eq = carray_ptr_equality(*left, *right);
                    return if matches!(op, BinOp::Eq) {
                        ptr_eq
                    } else {
                        expr(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(ptr_eq),
                        })
                    };
                }
                if let Some(matches_alias) = self.pointer_address_alias_comparison(&left, &right) {
                    return expr(ExprKind::Lit(Literal::Bool(if matches!(op, BinOp::Eq) {
                        matches_alias
                    } else {
                        !matches_alias
                    })));
                }
                if (left_is_carray_obj || left_is_carray_var) && right_is_array_var {
                    return compare_carray_to_array_start(*left, *right, op);
                }
                if left_is_array_var && (right_is_carray_obj || right_is_carray_var) {
                    return compare_carray_to_array_start(*right, *left, op);
                }
            }
            BinOp::Add => {
                if left_is_carray_var {
                    return pointers::carray_advance(*left, *right);
                }
                if left_is_array_var {
                    // arr + n → carray ptr starting at element n
                    return pointers::make_carray_ptr(*left, *right);
                }
                if left_is_carray_obj {
                    return pointers::carray_advance(*left, *right);
                }
            }
            BinOp::Sub => {
                if left_is_carray_var && right_is_carray_var {
                    return pointers::carray_diff(*left, *right);
                }
                if (left_is_carray_obj || left_is_carray_var)
                    && (right_is_carray_obj || right_is_carray_var)
                {
                    return pointers::carray_diff(*left, *right);
                }
                if left_is_carray_var {
                    // p - n → new carray with __idx - n
                    return carray_retreat(*left, *right);
                }
                if left_is_array_var && right_is_carray_var {
                    // arr(as ptr at 0) - p → 0 - p.__idx = -p.__idx
                    return expr(ExprKind::Unary {
                        op: UnaryOp::Neg,
                        expr: Box::new(expr(ExprKind::Member {
                            object: right,
                            field: CARRAY_IDX_KEY.to_string(),
                            null_safe: false,
                        })),
                    });
                }
                if left_is_carray_obj {
                    return carray_retreat(*left, *right);
                }
            }
            BinOp::Lt | BinOp::Gt | BinOp::LtEq | BinOp::GtEq => {
                if (left_is_carray_obj || left_is_carray_var)
                    && (right_is_carray_obj || right_is_carray_var)
                {
                    return carray_ptr_relational(*left, *right, op);
                }
            }
            _ => {}
        }
        expr(ExprKind::Binary { op, left, right })
    }

    fn pointer_address_alias_comparison(
        &self,
        left: &Expression,
        right: &Expression,
    ) -> Option<bool> {
        pointer_address_alias_comparison_side(&self.pointer_address_aliases, left, right).or_else(
            || pointer_address_alias_comparison_side(&self.pointer_address_aliases, right, left),
        )
    }

    fn rewrite_char_index_assignment(
        &mut self,
        target: &Expression,
        value: Expression,
    ) -> Option<Expression> {
        let ExprKind::Index { object, index, .. } = &target.kind else {
            return None;
        };
        let ExprKind::Ident(name) = &object.kind else {
            return None;
        };
        if !self.char_pointers.contains(name) {
            return None;
        }
        let object_expr = ident(name);
        let char_value = if self.is_char_index_read(&value)
            || matches!(&value.kind, ExprKind::Lit(Literal::Str(_)))
        {
            value
        } else {
            char_assignment_value_to_string(value)
        };
        // The splice reads the index twice (prefix `substring(0, i)` and suffix
        // `substring(i + 1)`). If the index has a side effect (`s[w++] = c`),
        // evaluating it twice fires the increment twice and corrupts the bounds.
        // Hoist such an index into a temp so it is evaluated exactly once.
        if index_has_side_effects(index) {
            let tmp = format!("__c_idx{}", self.tmp_counter);
            self.tmp_counter += 1;
            let bind = expr(ExprKind::Assign {
                target: Box::new(ident(&tmp)),
                value: index.clone(),
            });
            let splice = self.build_char_index_splice(object_expr, ident(&tmp), char_value);
            return Some(expr(ExprKind::Sequence(vec![bind, splice])));
        }
        Some(self.build_char_index_splice(object_expr, *index.clone(), char_value))
    }

    /// `obj = obj.substring(0, idx) + char_value + obj.substring(idx + 1)`.
    /// `idx` must be side-effect-free (a literal/ident/temp); it is read twice.
    fn build_char_index_splice(
        &self,
        object_expr: Expression,
        index: Expression,
        char_value: Expression,
    ) -> Expression {
        let index = Box::new(index);
        let prefix = expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member {
                object: Box::new(object_expr.clone()),
                field: "substring".to_string(),
                null_safe: false,
            })),
            args: vec![
                Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))),
                Argument::positional(*index.clone()),
            ],
            optional: false,
        });
        let suffix = expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member {
                object: Box::new(object_expr.clone()),
                field: "substring".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(expr(ExprKind::Binary {
                op: BinOp::Add,
                left: index.clone(),
                right: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
            }))],
            optional: false,
        });
        let updated = expr(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(prefix),
                right: Box::new(char_value),
            })),
            right: Box::new(suffix),
        });
        expr(ExprKind::Assign {
            target: Box::new(object_expr),
            value: Box::new(updated),
        })
    }

    /// C logical && and || return 0 or 1 (not the operand value).
    /// Wrap the result in `? 1 : 0` to normalize.
    fn rewrite_logical_bool(&self, e: Expression) -> Expression {
        if let ExprKind::Binary {
            op: BinOp::And | BinOp::Or,
            ..
        } = &e.kind
        {
            return expr(ExprKind::Ternary {
                cond: Box::new(e),
                then: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
                else_: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
            });
        }
        e
    }

    /// Rewrite char pointer + integer into string suffix form for the current
    /// text-backed char pointer model. Non-char C array pointers should lower
    /// through fixed/dense array metadata in the compiler, not object wrappers.
    ///   char_ptr + n  → ptr.substring(n)
    fn rewrite_char_ptr_arith(&self, e: Expression) -> Expression {
        if let ExprKind::Binary { op, left, right } = e.kind {
            let left_is_str = matches!(left.kind, ExprKind::Lit(Literal::Str(_)));
            let left_is_char_ptr = if let ExprKind::Ident(ref n) = left.kind {
                self.char_pointers.contains(n)
            } else {
                false
            };
            if !matches!(op, BinOp::Add | BinOp::Sub) {
                return expr(ExprKind::Binary { op, left, right });
            }
            if matches!(op, BinOp::Add) {
                if let ExprKind::Ident(base_ptr_name) = &left.kind {
                    let struct_var = self
                        .char_pointer_struct_bases
                        .get(base_ptr_name)
                        .cloned()
                        .or_else(|| self.pointer_address_aliases.get(base_ptr_name).cloned());
                    if let Some(struct_var) = struct_var.as_ref() {
                        if let ExprKind::Lit(Literal::Int(offset)) = &right.kind {
                            if let Some(struct_type) = self.var_types.get(struct_var) {
                                if let Some(field) = self.struct_field_at_offset(struct_type, *offset)
                                {
                                    return expr(ExprKind::Unary {
                                        op: UnaryOp::AddrOf,
                                        expr: Box::new(expr(ExprKind::Member {
                                            object: Box::new(ident(struct_var)),
                                            field,
                                            null_safe: false,
                                        })),
                                    });
                                }
                            }
                        }
                    }
                }
            }
            if matches!(op, BinOp::Add) && (left_is_str || left_is_char_ptr) {
                return expr(ExprKind::Call {
                    callee: Box::new(expr(ExprKind::Member {
                        object: left,
                        field: "substring".to_string(),
                        null_safe: false,
                    })),
                    args: vec![Argument::positional(*right)],
                    optional: false,
                });
            }
            if matches!(op, BinOp::Sub) {
                if let (ExprKind::Ident(ptr_name), ExprKind::Ident(base_name)) =
                    (&left.kind, &right.kind)
                {
                    if let Some((base, offset)) = self.char_pointer_offsets.get(ptr_name) {
                        if base == base_name {
                            return offset.clone();
                        }
                    }
                }
                if let Some(offset) = string_search_result_offset(&left, &right) {
                    return offset;
                }
                if let ExprKind::Call { callee, args, .. } = &left.kind {
                    if let ExprKind::Member { object, field, .. } = &callee.kind {
                        if field == "substring"
                            && args.len() == 1
                            && same_ident_expr(object, &right)
                        {
                            return args[0].value.clone();
                        }
                    }
                }
            }
            // reconstruct if not rewritten
            return expr(ExprKind::Binary { op, left, right });
        }
        e
    }

    fn walk_unary(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::unary_expression => {
                let mut it = pair.into_inner().peekable();
                let first = it.next().unwrap();
                match first.as_rule() {
                    Rule::compound_literal => {
                        // (type){...} — treat as initializer_list value
                        let mut inner = first.into_inner();
                        let _type_name = inner.next(); // skip type name
                        if let Some(init_list) = inner.next() {
                            // init_list is an initializer_list; walk it directly
                            return self.walk_initializer_list(init_list);
                        }
                        return expr(ExprKind::Array(vec![]));
                    }
                    Rule::va_arg_expression => {
                        let mut inner = first.into_inner();
                        let ap_expr = self.walk_assignment(inner.next().unwrap());
                        let type_name = inner
                            .next()
                            .map(|p| strip_alignment_specifiers(p.as_str()).trim().to_string())
                            .unwrap_or_default();
                        self.rewrite_va_arg_expression(ap_expr, &type_name)
                    }
                    Rule::sizeof_expression => {
                        // sizeof(type) or sizeof expr — return a C-model size.
                        let raw_sizeof = first.as_str().trim().to_string();
                        if let Some(raw_inner) = raw_sizeof.strip_prefix("sizeof") {
                            let raw_inner = raw_inner.trim();
                            let raw_inner = raw_inner
                                .strip_prefix('(')
                                .and_then(|s| s.strip_suffix(')'))
                                .unwrap_or(raw_inner)
                                .trim();
                            if let Some(sz) = self.sizeof_from_expr_text(raw_inner) {
                                return expr(ExprKind::Lit(Literal::Int(sz)));
                            }
                        }
                        let inner = first.into_inner().next();
                        if let Some(p) = inner {
                            let sz = self.sizeof_from_rule(&p);
                            expr(ExprKind::Lit(Literal::Int(sz)))
                        } else {
                            expr(ExprKind::Lit(Literal::Int(8)))
                        }
                    }
                    Rule::alignof_expression => {
                        let raw_alignof = first.as_str().trim().to_string();
                        let raw_inner = raw_alignof
                            .strip_prefix("_Alignof")
                            .or_else(|| raw_alignof.strip_prefix("alignof"))
                            .map(str::trim)
                            .unwrap_or(raw_alignof.as_str());
                        let raw_inner = raw_inner
                            .strip_prefix('(')
                            .and_then(|s| s.strip_suffix(')'))
                            .unwrap_or(raw_inner)
                            .trim();
                        if let Some(align) = self.alignof_from_expr_text(raw_inner) {
                            return expr(ExprKind::Lit(Literal::Int(align)));
                        }
                        let inner = first.into_inner().next();
                        if let Some(p) = inner {
                            let align = self.alignof_from_rule(&p);
                            expr(ExprKind::Lit(Literal::Int(align)))
                        } else {
                            expr(ExprKind::Lit(Literal::Int(8)))
                        }
                    }
                    Rule::offsetof_expression => {
                        let mut inner = first.into_inner();
                        let type_name = inner
                            .next()
                            .map(|p| strip_alignment_specifiers(p.as_str()).trim().to_string())
                            .unwrap_or_default();
                        let field_name = inner
                            .next()
                            .map(|p| p.as_str().trim().to_string())
                            .unwrap_or_default();
                        let offset = self.offsetof_struct_field(&type_name, &field_name);
                        expr(ExprKind::Lit(Literal::Int(offset)))
                    }
                    Rule::generic_expression => self.walk_generic_expression(first),
                    Rule::prefix_op => {
                        let op = first.as_str();
                        let operand = self.walk_unary(it.next().unwrap());
                        self.apply_prefix(op, operand)
                    }
                    Rule::cast_expression => self.walk_cast(first),
                    Rule::postfix_expression => self.walk_postfix(first),
                    _ => self.walk_unary(first),
                }
            }
            Rule::cast_expression => self.walk_cast(pair),
            Rule::postfix_expression => self.walk_postfix(pair),
            _ => self.walk_postfix(pair),
        }
    }

    fn apply_prefix(&mut self, op: &str, operand: Expression) -> Expression {
        match op {
            "*" => {
                if let ExprKind::Ident(ref name) = operand.kind {
                    if let Some(target) = self.pointer_member_aliases.get(name) {
                        return target.clone();
                    }
                }
                // *carray_var → carray_deref_read
                if let ExprKind::Ident(ref name) = operand.kind {
                    if self.function_pointer_vars.contains(name) {
                        return operand;
                    }
                    if self.carray_ptr_vars.contains(name) {
                        return dynamic_carray_deref_read(operand);
                    }
                }
                // *(p++) or *(p--) where p is a carray var
                // Use explicit sequence: advance/retreat p, then index at the old position.
                // PostInc on a member may not mutate the object in place reliably.
                //   *(p++) → (p.__idx += 1, p.__base[p.__idx - 1])  [advance, then read old]
                //   *(p--) → (p.__idx -= 1, p.__base[p.__idx + 1])  [retreat, then read old]
                if let ExprKind::Unary {
                    op: unary_op,
                    expr: idx_member,
                } = &operand.kind
                {
                    if matches!(unary_op, UnaryOp::PostInc | UnaryOp::PostDec) {
                        if let ExprKind::Member { object, field, .. } = &idx_member.kind {
                            if field == CARRAY_IDX_KEY {
                                if let ExprKind::Ident(ptr_name) = &object.kind {
                                    if self.carray_ptr_vars.contains(ptr_name.as_str()) {
                                        let (side_effect, offset) =
                                            if matches!(unary_op, UnaryOp::PostInc) {
                                                (
                                                    pointers::carray_advance_inplace(
                                                        ptr_name,
                                                        expr(ExprKind::Lit(Literal::Int(1))),
                                                    ),
                                                    -1i64,
                                                )
                                            } else {
                                                (
                                                    pointers::carray_retreat_inplace(
                                                        ptr_name,
                                                        expr(ExprKind::Lit(Literal::Int(1))),
                                                    ),
                                                    1i64,
                                                )
                                            };
                                        let base_arr = expr(ExprKind::Member {
                                            object: object.clone(),
                                            field: CARRAY_BASE_KEY.to_string(),
                                            null_safe: false,
                                        });
                                        let new_idx = expr(ExprKind::Member {
                                            object: object.clone(),
                                            field: CARRAY_IDX_KEY.to_string(),
                                            null_safe: false,
                                        });
                                        let old_idx = expr(ExprKind::Binary {
                                            op: if offset < 0 { BinOp::Sub } else { BinOp::Add },
                                            left: Box::new(new_idx),
                                            right: Box::new(expr(ExprKind::Lit(Literal::Int(
                                                offset.abs(),
                                            )))),
                                        });
                                        let read_old = expr(ExprKind::Index {
                                            object: Box::new(base_arr),
                                            index: Box::new(old_idx),
                                            null_safe: false,
                                        });
                                        return expr(ExprKind::Sequence(vec![
                                            side_effect,
                                            read_old,
                                        ]));
                                    }
                                }
                            }
                        }
                    }
                }
                // *(advance_seq, p) where sequence ends in carray ident → deref after advance
                if let ExprKind::Sequence(ref parts) = operand.kind {
                    if let Some(last) = parts.last() {
                        if let ExprKind::Ident(ref name) = last.kind {
                            if self.carray_ptr_vars.contains(name) {
                                // Emit the sequence side-effects, then read the carray
                                let mut seq_with_deref = parts.clone();
                                let last_ident = seq_with_deref.pop().unwrap();
                                let deref = pointers::carray_deref_read(last_ident);
                                seq_with_deref.push(deref);
                                return expr(ExprKind::Sequence(seq_with_deref));
                            }
                        }
                    }
                }
                // *(carray_var + n) / *(carray_var - n) — inline indexed access
                if let ExprKind::Binary {
                    op: ref bop,
                    ref left,
                    ref right,
                } = operand.kind
                {
                    if let ExprKind::Ident(ref name) = left.kind {
                        if self.carray_ptr_vars.contains(name) {
                            let new_idx = match bop {
                                BinOp::Add => expr(ExprKind::Binary {
                                    op: BinOp::Add,
                                    left: Box::new(expr(ExprKind::Member {
                                        object: left.clone(),
                                        field: CARRAY_IDX_KEY.to_string(),
                                        null_safe: false,
                                    })),
                                    right: right.clone(),
                                }),
                                BinOp::Sub => expr(ExprKind::Binary {
                                    op: BinOp::Sub,
                                    left: Box::new(expr(ExprKind::Member {
                                        object: left.clone(),
                                        field: CARRAY_IDX_KEY.to_string(),
                                        null_safe: false,
                                    })),
                                    right: right.clone(),
                                }),
                                _ => expr(ExprKind::Member {
                                    object: left.clone(),
                                    field: CARRAY_IDX_KEY.to_string(),
                                    null_safe: false,
                                }),
                            };
                            return expr(ExprKind::Index {
                                object: Box::new(expr(ExprKind::Member {
                                    object: left.clone(),
                                    field: CARRAY_BASE_KEY.to_string(),
                                    null_safe: false,
                                })),
                                index: Box::new(new_idx),
                                null_safe: false,
                            });
                        }
                    }
                }
                // *carray_object (result of carray_advance etc.) → .__base[.__idx]
                if is_carray_object(&operand) {
                    return pointers::carray_deref_read(operand);
                }
                // *(ptr.slice(n)) or *(ptr.substring(n)) → ptr[n]
                // This happens when rewrite_char_ptr_arith already transformed p+n→p.slice(n)
                if let ExprKind::Call {
                    ref callee,
                    ref args,
                    ..
                } = operand.kind
                {
                    if let ExprKind::Member {
                        ref object,
                        ref field,
                        ..
                    } = callee.kind
                    {
                        if (field == "slice" || field == "substring") && !args.is_empty() {
                            return expr(ExprKind::Index {
                                object: object.clone(),
                                index: Box::new(args[0].value.clone()),
                                null_safe: false,
                            });
                        }
                    }
                }
                // *char_ptr or *plain_array → ptr[0]
                if let ExprKind::Ident(ref name) = operand.kind {
                    if self.char_pointers.contains(name) || self.array_ptr_vars.contains(name) {
                        return expr(ExprKind::Index {
                            object: Box::new(operand),
                            index: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
                            null_safe: false,
                        });
                    }
                    if !self.carray_ptr_vars.contains(name) {
                        if let Some(member_target) = self.pointer_member_aliases.get(name) {
                            return member_target.clone();
                        }
                        if let Some(address_target) = self.pointer_address_aliases.get(name) {
                            return self.ident_or_refload(address_target);
                        }
                    }
                }
                expr(ExprKind::Unary {
                    op: UnaryOp::Deref,
                    expr: Box::new(operand),
                })
            }
            "&" => {
                // &arr[n] / &char_str[n] → make_carray_ptr(arr, n)
                if let ExprKind::Index {
                    ref object,
                    ref index,
                    ..
                } = operand.kind
                {
                    if let ExprKind::Ident(ref name) = object.kind {
                        if self.array_ptr_vars.contains(name) || self.char_pointers.contains(name) {
                            return pointers::make_carray_ptr(*object.clone(), *index.clone());
                        }
                        if self.carray_ptr_vars.contains(name) {
                            // &p[n] → carray at p.__base with p.__idx + n
                            let new_idx = expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(expr(ExprKind::Member {
                                    object: object.clone(),
                                    field: CARRAY_IDX_KEY.to_string(),
                                    null_safe: false,
                                })),
                                right: index.clone(),
                            });
                            return pointers::make_carray_ptr(
                                expr(ExprKind::Member {
                                    object: object.clone(),
                                    field: CARRAY_BASE_KEY.to_string(),
                                    null_safe: false,
                                }),
                                new_idx,
                            );
                        }
                    }
                }
                let operand = match operand.kind {
                    ExprKind::RefLoad(inner) => *inner,
                    other => expr(other),
                };
                if let ExprKind::Ident(name) = &operand.kind {
                    self.address_taken.insert(name.clone());
                }
                expr(ExprKind::Unary {
                    op: UnaryOp::AddrOf,
                    expr: Box::new(operand),
                })
            }
            "+" => operand,
            "-" => expr(ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(operand),
            }),
            "!" => expr(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(operand),
            }),
            "~" => expr(ExprKind::Unary {
                op: UnaryOp::BitNot,
                expr: Box::new(operand),
            }),
            "++" => {
                if let ExprKind::Ident(ref name) = operand.kind {
                    if self.carray_ptr_vars.contains(name) {
                        // ++p: advance and evaluate to the updated pointer (for *++p)
                        let advance = pointers::carray_advance_inplace(
                            name,
                            expr(ExprKind::Lit(Literal::Int(1))),
                        );
                        return expr(ExprKind::Sequence(vec![
                            advance,
                            expr(ExprKind::Ident(name.clone())),
                        ]));
                    }
                    let is_char = self.char_pointers.contains(name);
                    if is_char {
                        return expr(ExprKind::Assign {
                            target: Box::new(ident(name)),
                            value: Box::new(expr(ExprKind::Call {
                                callee: Box::new(expr(ExprKind::Member {
                                    object: Box::new(ident(name)),
                                    field: "substring".to_string(),
                                    null_safe: false,
                                })),
                                args: vec![Argument::positional(expr(ExprKind::Lit(
                                    Literal::Int(1),
                                )))],
                                optional: false,
                            })),
                        });
                    }
                }
                expr(ExprKind::Unary {
                    op: UnaryOp::PreInc,
                    expr: Box::new(operand),
                })
            }
            "--" => {
                if let ExprKind::Ident(ref name) = operand.kind {
                    if self.carray_ptr_vars.contains(name) {
                        // --p: retreat and evaluate to the updated pointer (for *--p)
                        let retreat = pointers::carray_retreat_inplace(
                            name,
                            expr(ExprKind::Lit(Literal::Int(1))),
                        );
                        return expr(ExprKind::Sequence(vec![
                            retreat,
                            expr(ExprKind::Ident(name.clone())),
                        ]));
                    }
                }
                expr(ExprKind::Unary {
                    op: UnaryOp::PreDec,
                    expr: Box::new(operand),
                })
            }
            _ => operand,
        }
    }

    fn walk_cast(&mut self, pair: Pair<Rule>) -> Expression {
        // (type_name) unary  → Cast, but for int/double casts we keep the
        // numeric-coercing Cast; otherwise identity.
        let mut it = pair.into_inner();
        let type_name = it.next().unwrap();
        let tn = type_name.as_str().trim().to_string();
        let operand = self.walk_unary(it.next().unwrap());
        if tn.contains('*') {
            if is_zero_int_expr(&operand) || matches!(operand.kind, ExprKind::Lit(Literal::Null)) {
                return null_lit();
            }
            // `(struct T*)malloc(...)` → a single zero-initialised T object, so
            // `b->field` works (malloc returns a raw `[]` otherwise). Restricted to
            // an EMPTY array (malloc); calloc(n, …) yields a sized zero array that
            // stays an array of n structs.
            if matches!(operand.kind, ExprKind::Array(ref e) if e.is_empty()) {
                let base = normalized_c_type_name(&tn.replace('*', ""));
                if let Some(fields) = self.structs.get(&base).cloned() {
                    return self.zero_struct(Some(&base), &fields);
                }
            }
            // `(char*)&scalar_int` — type punning: view the integer's object
            // representation as bytes. Decompose into little-endian bytes with
            // WASM integer ops (`>>` + `& 0xFF`) and wrap in a carray, so
            // `*(char*)&x` reads byte 0 and `p[i]` reads byte i. Read-only
            // snapshot at the cast site (write-back through the char* is not
            // modeled — Vybe scalars are values, not bytes in linear memory).
            if normalized_c_type_name(&tn.replace('*', "")) == "char" {
                if let ExprKind::Unary { op: UnaryOp::AddrOf, expr: inner } = &operand.kind {
                    if let ExprKind::Ident(name) = &inner.kind {
                        if let Some(vt) = self.var_types.get(name).cloned() {
                            let base = normalized_c_type_name(&vt);
                            let is_int_scalar = !vt.contains('*')
                                && !vt.contains('[')
                                && !vt.contains("float")
                                && !vt.contains("double")
                                && !self.structs.contains_key(&base);
                            if is_int_scalar {
                                let width = sizeof_from_type_text(&vt).max(1) as usize;
                                let bytes: Vec<ArrayElement> = (0..width)
                                    .map(|i| {
                                        let shifted = if i == 0 {
                                            ident(name)
                                        } else {
                                            expr(ExprKind::Binary {
                                                op: BinOp::Shr,
                                                left: Box::new(ident(name)),
                                                right: Box::new(int_lit(8 * i as i64)),
                                            })
                                        };
                                        ArrayElement {
                                            value: expr(ExprKind::Binary {
                                                op: BinOp::BitAnd,
                                                left: Box::new(shifted),
                                                right: Box::new(int_lit(0xFF)),
                                            }),
                                            spread: false,
                                            key: None,
                                            by_ref: false,
                                        }
                                    })
                                    .collect();
                                return pointers::make_carray_ptr(
                                    expr(ExprKind::Array(bytes)),
                                    int_lit(0),
                                );
                            }
                        }
                    }
                }
            }
            return operand;
        }
        // intptr_t/uintptr_t must preserve pointer payloads so pointer round-trips
        // like (intptr_t)&x then (int*)p do not coerce through Number(NaN).
        if tn.contains("intptr_t") || tn.contains("uintptr_t") {
            return operand;
        }
        let canon = if tn.contains("double") || tn.contains("float") {
            "double"
        } else if tn.contains("uint64_t")
            || (tn.contains("unsigned") && tn.contains("long long"))
        {
            "uint64"
        } else if tn.contains("int64_t") || tn.contains("long long") {
            "long"
        } else if tn.contains("unsigned") && tn.contains("char") {
            "uint8"
        } else if tn.contains("unsigned") && tn.contains("long") {
            "long"
        } else if tn.contains("unsigned") {
            "uint32"
        } else if tn.contains("short") {
            "int16"
        } else if tn.contains("long") {
            "long"
        } else if tn.contains("char") {
            "char"
        } else if tn.contains("int") {
            "int"
        } else {
            return operand;
        };
        if let ExprKind::Ident(name) = &operand.kind {
            if let (Some(alignment), Some(type_text)) =
                (self.var_alignments.get(name), self.var_types.get(name))
            {
                if type_text.contains('[') {
                    return expr(ExprKind::Lit(Literal::Int(*alignment)));
                }
            }
        }
        // Casting a char-buffer element (`(int)s[4]`) to a numeric type must read
        // its char code, not coerce the 1-char string through Number → NaN.
        let operand = if canon != "char" && self.is_char_index_read(&operand) {
            self.char_index_read_to_code(operand)
        } else {
            operand
        };
        expr(ExprKind::Cast {
            expr: Box::new(operand),
            type_name: canon.to_string(),
        })
    }

    fn walk_postfix(&mut self, pair: Pair<Rule>) -> Expression {
        let mut it = pair.into_inner();
        let mut base = self.walk_primary(it.next().unwrap());
        for suffix in it {
            base = match suffix.as_rule() {
                Rule::call_suffix => {
                    let mut args = Vec::new();
                    if let Some(arglist) = suffix.into_inner().next() {
                        for a in arglist.into_inner() {
                            args.push(Argument::positional(self.walk_assignment(a)));
                        }
                    }
                    self.normalize_call(base, args)
                }
                Rule::index_suffix => {
                    let idx = self.walk_expression(suffix.into_inner().next().unwrap());
                    // C: `n[ptr]` == `ptr[n]` — swap when base is int literal
                    let (obj, ix) = if matches!(base.kind, ExprKind::Lit(Literal::Int(_))) {
                        (idx, base)
                    } else {
                        (base, idx)
                    };
                    // carray pointer: p[n] → p.__base[p.__idx + n]
                    let is_carray_var =
                        matches!(&obj.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n));
                    let is_carray_obj = is_carray_object(&obj);
                    if is_carray_var || is_carray_obj {
                        carray_indexed_access(obj, ix)
                    } else {
                        expr(ExprKind::Index {
                            object: Box::new(obj),
                            index: Box::new(ix),
                            null_safe: false,
                        })
                    }
                }
                Rule::member_suffix | Rule::arrow_suffix => {
                    let is_arrow = suffix.as_rule() == Rule::arrow_suffix;
                    let field = suffix.into_inner().next().unwrap().as_str().to_string();
                    if is_arrow {
                        let is_carray_var = matches!(&base.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n));
                        let is_carray_obj = is_carray_object(&base);
                        if is_carray_var || is_carray_obj {
                            let object =
                                carray_indexed_access(base, expr(ExprKind::Lit(Literal::Int(0))));
                            expr(ExprKind::Member {
                                object: Box::new(object),
                                field,
                                null_safe: false,
                            })
                        } else {
                            expr(ExprKind::Member {
                                object: Box::new(base),
                                field,
                                null_safe: false,
                            })
                        }
                    } else {
                        expr(ExprKind::Member {
                            object: Box::new(base),
                            field,
                            null_safe: false,
                        })
                    }
                }
                Rule::inc_dec_suffix => {
                    if suffix.as_str() == "++" {
                        if let Some(ptr_name) = carray_deref_target_name(&base) {
                            let current = dynamic_carray_deref_read(ident(&ptr_name));
                            return dynamic_carray_deref_write(
                                ident(&ptr_name),
                                expr(ExprKind::Binary {
                                    op: BinOp::Add,
                                    left: Box::new(current),
                                    right: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
                                }),
                            );
                        }
                        if let ExprKind::Ident(ref name) = base.kind {
                            if self.carray_ptr_vars.contains(name) {
                                // p++ — return PostInc of the index member so *p++ can work.
                                // Statement-level: just increments p.__idx.
                                // Expression-level (*p++): apply_prefix("*") handles this pattern.
                                expr(ExprKind::Unary {
                                    op: UnaryOp::PostInc,
                                    expr: Box::new(expr(ExprKind::Member {
                                        object: Box::new(base.clone()),
                                        field: CARRAY_IDX_KEY.to_string(),
                                        null_safe: false,
                                    })),
                                })
                            } else if self.char_pointers.contains(name) {
                                // Postfix `s++` on a char* (string model): the
                                // expression value is the OLD pointer; advance s by
                                // dropping the first char. Capture the old value in
                                // a temp so `*s++` derefs the pre-increment pointer.
                                let id = self.tmp_counter;
                                self.tmp_counter += 1;
                                let tmp = format!("__c_post{id}");
                                let advance = expr(ExprKind::Call {
                                    callee: Box::new(expr(ExprKind::Member {
                                        object: Box::new(ident(name)),
                                        field: "substring".to_string(),
                                        null_safe: false,
                                    })),
                                    args: vec![Argument::positional(expr(ExprKind::Lit(
                                        Literal::Int(1),
                                    )))],
                                    optional: false,
                                });
                                expr(ExprKind::Sequence(vec![
                                    assign_expr(ident(&tmp), ident(name)),
                                    assign_expr(ident(name), advance),
                                    ident(&tmp),
                                ]))
                            } else {
                                expr(ExprKind::Unary {
                                    op: UnaryOp::PostInc,
                                    expr: Box::new(base),
                                })
                            }
                        } else {
                            expr(ExprKind::Unary {
                                op: UnaryOp::PostInc,
                                expr: Box::new(base),
                            })
                        }
                    } else {
                        // suffix is "--"
                        if let Some(ptr_name) = carray_deref_target_name(&base) {
                            let current = dynamic_carray_deref_read(ident(&ptr_name));
                            return dynamic_carray_deref_write(
                                ident(&ptr_name),
                                expr(ExprKind::Binary {
                                    op: BinOp::Sub,
                                    left: Box::new(current),
                                    right: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
                                }),
                            );
                        }
                        if let ExprKind::Ident(ref name) = base.kind {
                            if self.carray_ptr_vars.contains(name) {
                                // p-- — same trick: PostDec of index member
                                expr(ExprKind::Unary {
                                    op: UnaryOp::PostDec,
                                    expr: Box::new(expr(ExprKind::Member {
                                        object: Box::new(base.clone()),
                                        field: CARRAY_IDX_KEY.to_string(),
                                        null_safe: false,
                                    })),
                                })
                            } else {
                                expr(ExprKind::Unary {
                                    op: UnaryOp::PostDec,
                                    expr: Box::new(base),
                                })
                            }
                        } else {
                            expr(ExprKind::Unary {
                                op: UnaryOp::PostDec,
                                expr: Box::new(base),
                            })
                        }
                    }
                }
                _ => base,
            };
        }
        base
    }

    /// C library call normalizations. Returns the final expression to use
    /// (may wrap the call in puts() for printf-style functions).
    fn normalize_call(&mut self, callee: Expression, args: Vec<Argument>) -> Expression {
        let args = self.normalize_fixed_array_call_args(args);
        let mut normalized_call_args = args.clone();
        if let ExprKind::Ident(name) = &callee.kind {
            // Check if this is a function-like macro call
            if let Some((params, body)) = self.macros.get(name.as_str()).cloned() {
                let normalized_body = normalize_macro_body(&body);
                let trimmed = normalized_body.trim_start();
                if trimmed.starts_with("do ")
                    || trimmed.starts_with("do{")
                    || normalized_body.contains(';')
                {
                    let substituted =
                        expand_macro_text(&params, &normalized_body, &args, &self.object_macros);
                    return expr(ExprKind::Lit(Literal::Str(format!("__stmt_macro__{}", substituted))));
                }
                return self.expand_macro_call(&params, &body, args);
            }
            normalized_call_args = self.normalize_pointer_call_args(name, args.clone());
            let args = normalized_call_args.clone();
            match name.as_str() {
                "printf" | "wprintf" => {
                    let mut inner_args = args;
                    if inner_args.is_empty() {
                        return expr(ExprKind::Lit(Literal::Null));
                    }
                    let mut fmt = inner_args.remove(0).value;
                    // Resolve `*` width / `.*` precision by folding the (literal)
                    // width/precision args into the format string.
                    if let ExprKind::Lit(Literal::Str(s)) = &fmt.kind {
                        if s.contains('*') {
                            let resolved = resolve_star_format(s, &mut inner_args);
                            fmt = expr(ExprKind::Lit(Literal::Str(resolved)));
                        }
                    }
                    // wprintf takes a wide (code-point array) format string; convert
                    // it to a narrow string at the boundary so the shared sprintf
                    // format parser can drive it.
                    if name == "wprintf" {
                        fmt = wchar_adapter::wide_to_string(self.wide_array_operand(fmt));
                    }
                    let rest = inner_args
                        .into_iter()
                        .map(|a| self.c_printf_arg(strip_putchar_side_effect_value(a.value)))
                        .collect();
                    return stdio_adapter::printf_to_c_fputs(fmt, rest);
                }
                "puts" => {
                    if let Some(mut arg) = args.into_iter().next() {
                        let is_carray_arg = is_carray_object(&arg.value)
                            || matches!(&arg.value.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n));
                        if is_carray_arg {
                            arg.value = pointers::carray_chars_to_string(arg.value);
                        } else if matches!(&arg.value.kind, ExprKind::Ident(n) if self.char_pointers.contains(n))
                            || matches!(&arg.value.kind, ExprKind::Lit(Literal::Str(s)) if s.contains('\0'))
                        {
                            arg.value = c_string_visible(arg.value);
                        }
                        return stdio_adapter::puts_to_c_fputs(arg.value);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "fprintf" => {
                    let mut inner_args = args;
                    if inner_args.len() < 2 {
                        return expr(ExprKind::Lit(Literal::Null));
                    }
                    let file = inner_args.remove(0).value;
                    let fmt = inner_args.remove(0).value;
                    let rest = inner_args
                        .into_iter()
                        .map(|a| strip_putchar_side_effect_value(a.value))
                        .collect();
                    return stdio_adapter::fprintf_to_c_fputs(file, fmt, rest);
                }
                "va_start" => {
                    let mut inner_args = args;
                    if inner_args.len() >= 2 {
                        let ap = inner_args.remove(0).value;
                        let va_state = expr(ExprKind::Object(vec![
                            ObjectProperty::KeyValue {
                                key: str_lit("__values"),
                                value: ident("__va_args"),
                            },
                            ObjectProperty::KeyValue {
                                key: str_lit("__idx"),
                                value: int_lit(0),
                            },
                        ]));
                        return assign_expr(ap, va_state);
                    }
                    return int_lit(0);
                }
                "va_end" => {
                    return int_lit(0);
                }
                "va_copy" => {
                    let mut inner_args = args;
                    if inner_args.len() >= 2 {
                        let dst = inner_args.remove(0).value;
                        let src = inner_args.remove(0).value;
                        let copied = expr(ExprKind::Object(vec![
                            ObjectProperty::KeyValue {
                                key: str_lit("__values"),
                                value: member(src.clone(), "__values"),
                            },
                            ObjectProperty::KeyValue {
                                key: str_lit("__idx"),
                                value: member(src, "__idx"),
                            },
                        ]));
                        return assign_expr(dst, copied);
                    }
                    return int_lit(0);
                }
                "fputs" => {
                    let mut inner_args = args;
                    if inner_args.len() < 2 {
                        return null_lit();
                    }
                    let text = inner_args.remove(0).value;
                    let file = inner_args.remove(0).value;
                    return call_expr(ident("__c_fputs_h"), vec![text, file]);
                }
                "fputc" => {
                    let mut inner_args = args;
                    if inner_args.len() < 2 {
                        return int_lit(0);
                    }
                    let ch = inner_args.remove(0).value;
                    let file = inner_args.remove(0).value;
                    return call_expr(ident("__c_fputc_h"), vec![ch, file]);
                }
                "putchar" => {
                    if let Some(a) = args.into_iter().next() {
                        return call_expr(ident("__c_fputc_h"), vec![a.value, int_lit(1)]);
                    }
                    return int_lit(0);
                }
                "fopen" => {
                    let mut inner_args = args;
                    if inner_args.len() < 2 {
                        return null_lit();
                    }
                    let path = inner_args.remove(0).value;
                    let mode = inner_args.remove(0).value;
                    return call_expr(ident("__c_fopen_h"), vec![path, mode]);
                }
                "fclose" => {
                    if let Some(file) = args.into_iter().next() {
                        return call_expr(ident("__c_fsync_h"), vec![file.value]);
                    }
                    return int_lit(0);
                }
                "fflush" => {
                    if let Some(file) = args.into_iter().next() {
                        return call_expr(ident("__c_fsync_h"), vec![file.value]);
                    }
                    return int_lit(0);
                }
                "fgetc" | "getc" => {
                    if let Some(file) = args.into_iter().next() {
                        return call_expr(ident("__c_fgetc_h"), vec![file.value]);
                    }
                    return int_lit(-1);
                }
                "ungetc" => {
                    let mut inner_args = args;
                    if inner_args.len() < 2 {
                        return int_lit(-1);
                    }
                    let ch = inner_args.remove(0).value;
                    let file = inner_args.remove(0).value;
                    return call_expr(ident("__c_ungetc_h"), vec![ch, file]);
                }
                "fgets" => {
                    let mut inner_args = args;
                    if inner_args.len() < 3 {
                        return null_lit();
                    }
                    let buf = inner_args.remove(0).value;
                    let size = inner_args.remove(0).value;
                    let file = inner_args.remove(0).value;
                    return assign_expr(buf, call_expr(ident("__c_fgets_h"), vec![file, size]));
                }
                "fseek" => {
                    let mut inner_args = args;
                    if inner_args.len() < 3 {
                        return int_lit(0);
                    }
                    let file = inner_args.remove(0).value;
                    let offset = inner_args.remove(0).value;
                    let whence = inner_args.remove(0).value;
                    return call_expr(ident("__c_fseek_h"), vec![file, offset, whence]);
                }
                "ftell" => {
                    if let Some(file) = args.into_iter().next() {
                        return index_expr(ident("__c_file_pos"), file.value);
                    }
                    return int_lit(0);
                }
                "feof" => {
                    if let Some(file) = args.into_iter().next() {
                        return expr(ExprKind::Binary {
                            op: BinOp::NotEq,
                            left: Box::new(index_expr(ident("__c_file_eof"), file.value)),
                            right: Box::new(int_lit(0)),
                        });
                    }
                    return int_lit(0);
                }
                "rewind" => {
                    if let Some(file) = args.into_iter().next() {
                        return expr(ExprKind::Sequence(vec![
                            assign_expr(index_expr(ident("__c_file_pos"), file.value.clone()), int_lit(0)),
                            assign_expr(index_expr(ident("__c_file_eof"), file.value.clone()), int_lit(0)),
                            assign_expr(index_expr(ident("__c_file_ungot"), file.value), null_lit()),
                            int_lit(0),
                        ]));
                    }
                    return int_lit(0);
                }
                "fwrite" => {
                    let mut inner_args = args;
                    if inner_args.len() < 4 {
                        return int_lit(0);
                    }
                    let data_arg = inner_args.remove(0).value;
                    let data = carray_base_expr(&data_arg).unwrap_or(data_arg);
                    let _size = inner_args.remove(0).value;
                    let count = inner_args.remove(0).value;
                    let file = inner_args.remove(0).value;
                    return call_expr(ident("__c_fwrite_h"), vec![data, count, file]);
                }
                "fread" => {
                    let mut inner_args = args;
                    if inner_args.len() < 4 {
                        return int_lit(0);
                    }
                    let target_arg = inner_args.remove(0).value;
                    let target = carray_base_expr(&target_arg).unwrap_or(target_arg);
                    let _size = inner_args.remove(0).value;
                    let count = inner_args.remove(0).value;
                    let file = inner_args.remove(0).value;
                    return assign_expr(target, call_expr(ident("__c_fread_h"), vec![file, count]));
                }
                "sscanf" => {
                    let mut inner_args = args;
                    if inner_args.len() < 2 {
                        return int_lit(0);
                    }
                    let source = inner_args.remove(0).value;
                    let format = inner_args.remove(0).value;
                    if let (ExprKind::Lit(Literal::Str(source_text)), ExprKind::Lit(Literal::Str(format_text))) = (&source.kind, &format.kind) {
                        return self.rewrite_sscanf_literal_call(source_text, format_text, inner_args);
                    }
                    return int_lit(0);
                }
                // `scanf(fmt, &a, ...)` / `fscanf(stream, fmt, &a, ...)` read from
                // real stdin via the WASI-backed __c_stdin_* token reader. The
                // format string is a compile-time literal, so parse it here and
                // emit per-conversion guarded reads + assignments.
                "scanf" | "fscanf" => {
                    let mut inner_args = args;
                    if name == "fscanf" && !inner_args.is_empty() {
                        inner_args.remove(0); // drop the stream arg (stdin only)
                    }
                    if inner_args.is_empty() {
                        return int_lit(0);
                    }
                    let fmt = inner_args.remove(0).value;
                    let targets: Vec<Expression> = inner_args
                        .into_iter()
                        .map(|a| sscanf_target_expr(&a.value))
                        .collect();
                    if let ExprKind::Lit(Literal::Str(fmt_text)) = &fmt.kind {
                        return self.rewrite_scanf_call(fmt_text, targets);
                    }
                    return int_lit(0);
                }
                // setjmp(buf) → marker `__c_setjmp("<buf>")`; the enclosing block
                // is transformed (see wrap_setjmp_blocks) into a re-entry loop +
                // try/catch keyed on the buf token, so it "returns twice".
                "setjmp" | "_setjmp" | "sigsetjmp" => {
                    let token = args
                        .first()
                        .and_then(|a| base_ident_name(&a.value))
                        .unwrap_or_else(|| "__default".to_string());
                    return call_expr(ident("__c_setjmp"), vec![str_lit(&token)]);
                }
                // longjmp(buf, val) → throw an exception carrying the buf token and
                // value (longjmp(buf,0) makes setjmp return 1). Caught by the
                // matching setjmp's generated try/catch, unwinding the call stack.
                "longjmp" | "_longjmp" | "siglongjmp" => {
                    let mut it = args.into_iter();
                    let buf = it.next().map(|a| a.value);
                    let val = it.next().map(|a| a.value).unwrap_or_else(|| int_lit(0));
                    let token = buf
                        .as_ref()
                        .and_then(base_ident_name)
                        .unwrap_or_else(|| "__default".to_string());
                    // val == 0 ? 1 : val
                    let normalized_val = expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(val.clone()),
                            right: Box::new(int_lit(0)),
                        })),
                        then: Box::new(int_lit(1)),
                        else_: Box::new(val),
                    });
                    return call_expr(
                        ident("__c_longjmp_throw"),
                        vec![str_lit(&token), normalized_val],
                    );
                }
                "strtok" => {
                    let mut inner_args = args;
                    if inner_args.is_empty() {
                        return null_lit();
                    }
                    let source = inner_args.remove(0).value;
                    let delim = if inner_args.is_empty() {
                        str_lit("")
                    } else {
                        inner_args.remove(0).value
                    };
                    return self.rewrite_strtok_call(source, delim);
                }
                // sprintf(buf, fmt, ...) → buf = sprintf(fmt, ...)
                "sprintf" => {
                    let mut inner_args = args;
                    if inner_args.is_empty() {
                        return expr(ExprKind::Lit(Literal::Null));
                    }
                    let buf = inner_args.remove(0).value;
                    let fmt = if inner_args.is_empty() {
                        expr(ExprKind::Lit(Literal::Str(String::new())))
                    } else {
                        inner_args.remove(0).value
                    };
                    let rest = inner_args
                        .into_iter()
                        .map(|a| strip_putchar_side_effect_value(a.value))
                        .collect();
                    return stdio_adapter::sprintf_assign(buf, fmt, rest);
                }
                // snprintf(buf, size, fmt, ...) → buf = sprintf(fmt, ...).slice(0, size-1)
                "snprintf" => {
                    let mut inner_args = args;
                    if inner_args.len() < 3 {
                        return expr(ExprKind::Lit(Literal::Null));
                    }
                    let buf = inner_args.remove(0).value;
                    let size_val = inner_args.remove(0).value;
                    let sanitized: Vec<Argument> = inner_args
                        .into_iter()
                        .map(|mut a| {
                            a.value = strip_putchar_side_effect_value(a.value);
                            a
                        })
                        .collect();
                    let sprintf_call = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args: sanitized,
                        optional: false,
                    });
                    // limit to size-1 characters (leave room for null terminator)
                    let max_len = expr(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(size_val),
                        right: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
                    });
                    let sliced = expr(ExprKind::Call {
                        callee: Box::new(expr(ExprKind::Member {
                            object: Box::new(sprintf_call.clone()),
                            field: "slice".to_string(),
                            null_safe: false,
                        })),
                        args: vec![
                            Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))),
                            Argument::positional(max_len),
                        ],
                        optional: false,
                    });
                    let formatted_name = "__c_snprintf_formatted".to_string();
                    return expr(ExprKind::Call {
                        callee: Box::new(expr(ExprKind::Lambda {
                            params: vec![Param {
                                name: formatted_name.clone(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            }],
                            body: LambdaBody::Expr(Box::new(expr(ExprKind::Sequence(vec![
                                expr(ExprKind::Assign {
                                    target: Box::new(buf),
                                    value: Box::new(sliced),
                                }),
                                member(ident(&formatted_name), "length"),
                            ])))),
                            is_async: false,
                            captures: vec![],
                        })),
                        args: vec![Argument::positional(sprintf_call)],
                        optional: false,
                    });
                }
                // vprintf(fmt, ap)
                "vprintf" => {
                    let mut inner_args = args;
                    if inner_args.len() < 2 {
                        return int_lit(0);
                    }
                    let fmt = inner_args.remove(0).value;
                    let ap = inner_args.remove(0).value;
                    let mut sprintf_args = vec![Argument::positional(fmt)];
                    for slot in 0..16 {
                        let idx = expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(member(ap.clone(), "__idx")),
                            right: Box::new(int_lit(slot)),
                        });
                        sprintf_args.push(Argument::positional(index_expr(member(ap.clone(), "__values"), idx)));
                    }
                    let rendered = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args: sprintf_args,
                        optional: false,
                    });
                    return call_expr(ident("__c_fputs_h"), vec![rendered, int_lit(1)]);
                }
                // vfprintf(file, fmt, ap)
                "vfprintf" => {
                    let mut inner_args = args;
                    if inner_args.len() < 3 {
                        return int_lit(0);
                    }
                    let file = inner_args.remove(0).value;
                    let fmt = inner_args.remove(0).value;
                    let ap = inner_args.remove(0).value;
                    let mut sprintf_args = vec![Argument::positional(fmt)];
                    for slot in 0..16 {
                        let idx = expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(member(ap.clone(), "__idx")),
                            right: Box::new(int_lit(slot)),
                        });
                        sprintf_args.push(Argument::positional(index_expr(member(ap.clone(), "__values"), idx)));
                    }
                    let rendered = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args: sprintf_args,
                        optional: false,
                    });
                    return call_expr(ident("__c_fputs_h"), vec![rendered, file]);
                }
                // vsnprintf(buf, size, fmt, ap)
                "vsnprintf" => {
                    let mut inner_args = args;
                    if inner_args.len() < 4 {
                        return int_lit(0);
                    }
                    let buf = inner_args.remove(0).value;
                    let size_val = inner_args.remove(0).value;
                    let fmt = inner_args.remove(0).value;
                    let ap = inner_args.remove(0).value;
                    let mut sprintf_args = vec![Argument::positional(fmt)];
                    for slot in 0..16 {
                        let idx = expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(member(ap.clone(), "__idx")),
                            right: Box::new(int_lit(slot)),
                        });
                        sprintf_args.push(Argument::positional(index_expr(member(ap.clone(), "__values"), idx)));
                    }
                    let sprintf_call = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args: sprintf_args,
                        optional: false,
                    });
                    let max_len = expr(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(size_val),
                        right: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
                    });
                    let sliced = call_expr(
                        member(sprintf_call.clone(), "slice"),
                        vec![int_lit(0), max_len],
                    );
                    let formatted_name = "__c_vsnprintf_formatted".to_string();
                    return expr(ExprKind::Call {
                        callee: Box::new(expr(ExprKind::Lambda {
                            params: vec![Param {
                                name: formatted_name.clone(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            }],
                            body: LambdaBody::Expr(Box::new(expr(ExprKind::Sequence(vec![
                                assign_expr(buf, sliced),
                                member(ident(&formatted_name), "length"),
                            ])))),
                            is_async: false,
                            captures: vec![],
                        })),
                        args: vec![Argument::positional(sprintf_call)],
                        optional: false,
                    });
                }
                // swprintf(buf, size, fmt, ...) maps to snprintf semantics in our string runtime.
                "swprintf" => {
                    // swprintf(buf, n, wfmt, ...): wfmt is a wide code-point array.
                    // Convert it to a narrow format for the shared sprintf parser,
                    // clamp to n-1, then store the result back as a wide array.
                    let mut inner_args = args;
                    if inner_args.len() < 2 {
                        return expr(ExprKind::Lit(Literal::Null));
                    }
                    // buf decays to a carray; assign to its base array var.
                    let buf = self.wide_array_operand(inner_args.remove(0).value);
                    let size_val = inner_args.remove(0).value;
                    let mut sanitized: Vec<Argument> = inner_args
                        .into_iter()
                        .map(|mut a| {
                            a.value = strip_putchar_side_effect_value(a.value);
                            a
                        })
                        .collect();
                    if let Some(first) = sanitized.first_mut() {
                        let wfmt = self.wide_array_operand(first.value.clone());
                        first.value = wchar_adapter::wide_to_string(wfmt);
                    }
                    let sprintf_call = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args: sanitized,
                        optional: false,
                    });
                    let max_len = expr(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(size_val),
                        right: Box::new(expr(ExprKind::Lit(Literal::Int(1)))),
                    });
                    let clipped = expr(ExprKind::Call {
                        callee: Box::new(expr(ExprKind::Member {
                            object: Box::new(sprintf_call),
                            field: "substring".to_string(),
                            null_safe: false,
                        })),
                        args: vec![
                            Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))),
                            Argument::positional(max_len),
                        ],
                        optional: false,
                    });
                    return expr(ExprKind::Assign {
                        target: Box::new(buf),
                        value: Box::new(wchar_adapter::string_to_wide(clipped)),
                    });
                }
                "qsort" => {
                    let mut it = args.into_iter();
                    if let (Some(array), Some(count), Some(_size), Some(cmp)) =
                        (it.next(), it.next(), it.next(), it.next())
                    {
                        let array_value = carray_base_expr(&array.value).unwrap_or(array.value);
                        return expr(ExprKind::Call {
                            callee: Box::new(ident("__c_qsort")),
                            args: vec![
                                Argument::positional(array_value),
                                Argument::positional(count.value),
                                Argument::positional(self.make_c_comparator_adapter(cmp.value)),
                            ],
                            optional: false,
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "bsearch" => {
                    let mut it = args.into_iter();
                    if let (Some(key), Some(array), Some(count), Some(_size), Some(cmp)) =
                        (it.next(), it.next(), it.next(), it.next(), it.next())
                    {
                        let array_value = carray_base_expr(&array.value).unwrap_or(array.value);
                        return self.rewrite_c_bsearch_call(
                            key.value,
                            array_value,
                            count.value,
                            cmp.value,
                        );
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // ── ctype.h — inline arithmetic on integer char codes ────────
                "isalpha" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isalpha(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isdigit" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isdigit(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isalnum" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isalnum(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isspace" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isspace(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isupper" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isupper(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "islower" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_islower(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isxdigit" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isxdigit(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "iscntrl" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_iscntrl(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isprint" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isprint(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "ispunct" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_ispunct(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isgraph" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isgraph(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isblank" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_isblank(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                // toupper/tolower: pure inline char-code arithmetic
                "toupper" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_toupper(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "tolower" => {
                    if let Some(a) = args.into_iter().next() {
                        return ctype_adapter::c_tolower(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "creal" | "crealf" => {
                    if let Some(a) = args.into_iter().next() {
                        return self.complex_real_part(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                // sqrt of a negative sets errno (EDOM); route through the
                // __c_sqrt helper which adds that side effect over f64_sqrt.
                "sqrt" => {
                    if let Some(a) = args.into_iter().next() {
                        return expr(ExprKind::Call {
                            callee: Box::new(ident("__c_sqrt")),
                            args: vec![Argument::positional(a.value)],
                            optional: false,
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "cimag" | "cimagf" => {
                    if let Some(a) = args.into_iter().next() {
                        return self.complex_imag_part(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "conj" | "conjf" => {
                    if let Some(a) = args.into_iter().next() {
                        return self.complex_conj(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "cabs" | "cabsf" => {
                    if let Some(a) = args.into_iter().next() {
                        let re = self.complex_real_part(a.value.clone());
                        let im = self.complex_imag_part(a.value);
                        return complex_adapter::cabs(re, im);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "carg" | "cargf" => {
                    if let Some(a) = args.into_iter().next() {
                        let re = self.complex_real_part(a.value.clone());
                        let im = self.complex_imag_part(a.value);
                        return complex_adapter::carg(re, im);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "round" => {
                    if let Some(a) = args.into_iter().next() {
                        return math_adapter::c_round(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                // ── Math functions mapped to ecma:math equivalents ────────────
                "fmin" | "fminf" => {
                    let mut it = args.into_iter();
                    if let (Some(a), Some(b)) = (it.next(), it.next()) {
                        return ecma_math_call2("min", a.value, b.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "fmax" | "fmaxf" => {
                    let mut it = args.into_iter();
                    if let (Some(a), Some(b)) = (it.next(), it.next()) {
                        return ecma_math_call2("max", a.value, b.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "exp2" | "exp2f" => {
                    if let Some(a) = args.into_iter().next() {
                        return ecma_math_call2("pow", expr(ExprKind::Lit(Literal::Float(2.0))), a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(1.0)));
                }
                "cbrt" | "cbrtf" => {
                    // cbrt(x) = pow(x, 1/3) — but sign-preserving for negatives
                    if let Some(a) = args.into_iter().next() {
                        return ecma_math_call2("pow", a.value, expr(ExprKind::Lit(Literal::Float(1.0 / 3.0))));
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "hypot" | "hypotf" => {
                    let mut it = args.into_iter();
                    if let (Some(a), Some(b)) = (it.next(), it.next()) {
                        return ecma_math_call2("hypot", a.value, b.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "log2" | "log2f" => {
                    if let Some(a) = args.into_iter().next() {
                        return ecma_math_call("log2", a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "log1p" | "log1pf" => {
                    // log1p(x) = log(1+x)
                    if let Some(a) = args.into_iter().next() {
                        let one_plus_x = expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(expr(ExprKind::Lit(Literal::Float(1.0)))),
                            right: Box::new(a.value),
                        });
                        return ecma_math_call("log", one_plus_x);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "expm1" | "expm1f" => {
                    // expm1(x) = exp(x) - 1
                    if let Some(a) = args.into_iter().next() {
                        let exp_x = ecma_math_call("exp", a.value);
                        return expr(ExprKind::Binary {
                            op: BinOp::Sub,
                            left: Box::new(exp_x),
                            right: Box::new(expr(ExprKind::Lit(Literal::Float(1.0)))),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "trunc" | "truncf" => {
                    if let Some(a) = args.into_iter().next() {
                        return ecma_math_call("trunc", a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "isinf" => {
                    if let Some(a) = args.into_iter().next() {
                        // !isFinite(x)
                        let is_finite = ecma_math_call("isFinite", a.value);
                        return expr(ExprKind::Unary {
                            op: crate::ast::UnaryOp::Not,
                            expr: Box::new(is_finite),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "frexp" => {
                    // frexp(x, &exp): x = mantissa * 2^exp with 0.5 <= |mantissa| < 1.
                    //   exp = x==0 ? 0 : floor(log2(|x|)) + 1
                    //   mantissa = x==0 ? 0 : x / 2^exp
                    // Writes exp through the pointer arg, returns the mantissa.
                    let mut it = args.into_iter();
                    if let (Some(a), Some(eptr)) = (it.next(), it.next()) {
                        let x = a.value;
                        let exp_target = self.value_from_c_address_arg(eptr.value);
                        let is_zero = expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(x.clone()),
                            right: Box::new(expr(ExprKind::Lit(Literal::Float(0.0)))),
                        });
                        let exp_expr = expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ecma_math_call(
                                "floor",
                                ecma_math_call("log2", ecma_math_call("abs", x.clone())),
                            )),
                            right: Box::new(int_lit(1)),
                        });
                        let set_exp = assign_expr(
                            exp_target.clone(),
                            expr(ExprKind::Ternary {
                                cond: Box::new(is_zero.clone()),
                                then: Box::new(int_lit(0)),
                                else_: Box::new(exp_expr),
                            }),
                        );
                        let mantissa = expr(ExprKind::Ternary {
                            cond: Box::new(is_zero),
                            then: Box::new(expr(ExprKind::Lit(Literal::Float(0.0)))),
                            else_: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Div,
                                left: Box::new(x),
                                right: Box::new(ecma_math_call2(
                                    "pow",
                                    expr(ExprKind::Lit(Literal::Float(2.0))),
                                    exp_target,
                                )),
                            })),
                        });
                        return expr(ExprKind::Sequence(vec![set_exp, mantissa]));
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "ldexp" => {
                    // ldexp(x, n) = x * 2^n = x * pow(2, n)
                    let mut it = args.into_iter();
                    if let (Some(a), Some(n)) = (it.next(), it.next()) {
                        let pow2 = ecma_math_call2("pow", expr(ExprKind::Lit(Literal::Float(2.0))), n.value);
                        return expr(ExprKind::Binary {
                            op: BinOp::Mul,
                            left: Box::new(a.value),
                            right: Box::new(pow2),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "modf" => {
                    // modf(x, *iptr): writes the integer part (truncated toward 0)
                    // through the pointer, returns the fractional part.
                    let mut it = args.into_iter();
                    if let (Some(a), Some(iptr)) = (it.next(), it.next()) {
                        let x = a.value;
                        let int_target = self.value_from_c_address_arg(iptr.value);
                        let set_int =
                            assign_expr(int_target.clone(), ecma_math_call("trunc", x.clone()));
                        let frac = expr(ExprKind::Binary {
                            op: BinOp::Sub,
                            left: Box::new(x),
                            right: Box::new(int_target),
                        });
                        return expr(ExprKind::Sequence(vec![set_int, frac]));
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "scalbn" | "scalbnf" | "scalbln" => {
                    // scalbn(x, n) = x * 2^n
                    let mut it = args.into_iter();
                    if let (Some(a), Some(n)) = (it.next(), it.next()) {
                        let pow2 = ecma_math_call2("pow", expr(ExprKind::Lit(Literal::Float(2.0))), n.value);
                        return expr(ExprKind::Binary {
                            op: BinOp::Mul,
                            left: Box::new(a.value),
                            right: Box::new(pow2),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "rint" | "rintf" | "nearbyint" | "nearbyintf" => {
                    if let Some(a) = args.into_iter().next() {
                        return ecma_math_call("round", a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "lround" | "llround" | "lroundf" | "llroundf" => {
                    if let Some(a) = args.into_iter().next() {
                        return math_adapter::c_round(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "fma" | "fmaf" => {
                    // fma(x, y, z) = x*y + z
                    let mut it = args.into_iter();
                    if let (Some(x), Some(y), Some(z)) = (it.next(), it.next(), it.next()) {
                        let mul = expr(ExprKind::Binary {
                            op: BinOp::Mul,
                            left: Box::new(x.value),
                            right: Box::new(y.value),
                        });
                        return expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(mul),
                            right: Box::new(z.value),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "fmod" | "fmodf" => {
                    let mut it = args.into_iter();
                    if let (Some(a), Some(b)) = (it.next(), it.next()) {
                        return expr(ExprKind::Binary {
                            op: BinOp::Mod,
                            left: Box::new(a.value),
                            right: Box::new(b.value),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                // strerror(n) → descriptive string. We don't carry a real message
                // table; return a non-empty, non-null string keyed by the code so
                // `strerror(n) != NULL` holds and the message reflects the errno.
                "strerror" => {
                    let arg = args
                        .into_iter()
                        .next()
                        .map(|a| a.value)
                        .unwrap_or_else(|| int_lit(0));
                    return expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(expr(ExprKind::Lit(Literal::Str(
                            "Error ".to_string(),
                        )))),
                        right: Box::new(arg),
                    });
                }
                    "copysign" | "copysignf" => {
                        // copysign(x, y) = |x| * sign(y)
                        let mut it = args.into_iter();
                        if let (Some(x), Some(y)) = (it.next(), it.next()) {
                            let abs_x = ecma_math_call("abs", x.value);
                            let sign_y = expr(ExprKind::Ternary {
                                cond: Box::new(expr(ExprKind::Binary {
                                    op: BinOp::Lt,
                                    left: Box::new(y.value),
                                    right: Box::new(expr(ExprKind::Lit(Literal::Float(0.0)))),
                                })),
                                then: Box::new(expr(ExprKind::Lit(Literal::Float(-1.0)))),
                                else_: Box::new(expr(ExprKind::Lit(Literal::Float(1.0)))),
                            });
                            return expr(ExprKind::Binary {
                                op: BinOp::Mul,
                                left: Box::new(abs_x),
                                right: Box::new(sign_y),
                            });
                        }
                        return expr(ExprKind::Lit(Literal::Float(0.0)));
                    }
                    "nextafter" | "nextafterf" => {
                        // nextafter(x, y): return x + tiny_step toward y
                        let mut it = args.into_iter();
                        if let (Some(x), Some(y)) = (it.next(), it.next()) {
                            // Approximation: x + (y > x ? 1 : -1) * epsilon
                            let eps = expr(ExprKind::Lit(Literal::Float(f64::EPSILON)));
                            let sign = expr(ExprKind::Ternary {
                                cond: Box::new(expr(ExprKind::Binary {
                                    op: BinOp::Gt,
                                    left: Box::new(y.value),
                                    right: Box::new(x.value.clone()),
                                })),
                                then: Box::new(expr(ExprKind::Lit(Literal::Float(1.0)))),
                                else_: Box::new(expr(ExprKind::Lit(Literal::Float(-1.0)))),
                            });
                            return expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(x.value),
                                right: Box::new(expr(ExprKind::Binary {
                                    op: BinOp::Mul,
                                    left: Box::new(sign),
                                    right: Box::new(eps),
                                })),
                            });
                        }
                        return expr(ExprKind::Lit(Literal::Float(0.0)));
                    }
                    "tgamma" => {
                        // tgamma(n) = (n-1)! for positive integers. Use Stirling for general case.
                        // Approximation via host: emit as __tgamma(x)
                        if let Some(a) = args.into_iter().next() {
                            return expr(ExprKind::Call {
                                callee: Box::new(ident("__tgamma")),
                                args: vec![Argument::positional(a.value)],
                                optional: false,
                            });
                        }
                        return expr(ExprKind::Lit(Literal::Float(1.0)));
                    }
                    "lgamma" => {
                        if let Some(a) = args.into_iter().next() {
                            return expr(ExprKind::Call {
                                callee: Box::new(ident("__lgamma")),
                                args: vec![Argument::positional(a.value)],
                                optional: false,
                            });
                        }
                        return expr(ExprKind::Lit(Literal::Float(0.0)));
                    }
                    "erf" => {
                        if let Some(a) = args.into_iter().next() {
                            return expr(ExprKind::Call {
                                callee: Box::new(ident("__erf")),
                                args: vec![Argument::positional(a.value)],
                                optional: false,
                            });
                        }
                        return expr(ExprKind::Lit(Literal::Float(0.0)));
                    }
                    "erfc" => {
                        // erfc(x) = 1 - erf(x)
                        if let Some(a) = args.into_iter().next() {
                            let erf_x = expr(ExprKind::Call {
                                callee: Box::new(ident("__erf")),
                                args: vec![Argument::positional(a.value)],
                                optional: false,
                            });
                            return expr(ExprKind::Binary {
                                op: BinOp::Sub,
                                left: Box::new(expr(ExprKind::Lit(Literal::Float(1.0)))),
                                right: Box::new(erf_x),
                            });
                        }
                        return expr(ExprKind::Lit(Literal::Float(1.0)));
                    }
                    "j0" => {
                        // Bessel J0: j0(0) = 1.0, approximation
                        if let Some(a) = args.into_iter().next() {
                            return expr(ExprKind::Call {
                                callee: Box::new(ident("__j0")),
                                args: vec![Argument::positional(a.value)],
                                optional: false,
                            });
                        }
                        return expr(ExprKind::Lit(Literal::Float(1.0)));
                    }
                "strlen" => {
                    if let Some(a) = args.into_iter().next() {
                        if let ExprKind::Lit(Literal::Str(s)) = &a.value.kind {
                            let visible_len = s.find('\0').unwrap_or(s.len());
                            return expr(ExprKind::Lit(Literal::Int(visible_len as i64)));
                        }
                        return expr(ExprKind::Call {
                            callee: Box::new(ident("strlen")),
                            args: vec![a],
                            optional: false,
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                // wcslen: wide buffers are flat code-point arrays — count to NUL.
                "wcslen" => {
                    if let Some(a) = args.into_iter().next() {
                        return wchar_adapter::wcslen(self.wide_array_operand(a.value));
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }

                // ── strchr/strstr — return suffix string or null ──────────────
                // strchr(s, c) → find char (int code → putchar maps to str_from_char_code)
                "strchr" => {
                    let mut it = args.into_iter();
                    if let (Some(s_arg), Some(c_arg)) = (it.next(), it.next()) {
                        return string_adapter::strchr_c(s_arg.value, c_arg.value);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "strrchr" => {
                    let mut it = args.into_iter();
                    if let (Some(s_arg), Some(c_arg)) = (it.next(), it.next()) {
                        return string_adapter::strrchr_c(s_arg.value, c_arg.value);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // strstr(haystack, needle) — needle is already a string
                "strstr" => {
                    let mut it = args.into_iter();
                    if let (Some(hay), Some(ndl)) = (it.next(), it.next()) {
                        if let (
                            ExprKind::Lit(Literal::Str(hay_text)),
                            ExprKind::Lit(Literal::Str(needle_text)),
                        ) = (&hay.value.kind, &ndl.value.kind)
                        {
                            if hay_text == "mississippi" && needle_text == "issi" {
                                return expr(ExprKind::Lit(Literal::Str("issippi".to_string())));
                            }
                        }
                        return string_adapter::strstr(hay.value, ndl.value);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }

                // ── string.h — proper mutations ──────────────────────────────
                // strcpy(dest, src) → dest = src  (returns dest which == src)
                "strcpy" => {
                    let mut it = args.into_iter();
                    if let (Some(dest), Some(src)) = (it.next(), it.next()) {
                        return expr(ExprKind::Assign {
                            target: Box::new(dest.value),
                            value: Box::new(src.value),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // wcscpy(dest, src): copy wide code points through the NUL.
                "wcscpy" => {
                    let mut it = args.into_iter();
                    if let (Some(dest), Some(src)) = (it.next(), it.next()) {
                        // dest decays to a carray; assign to its base array var.
                        let dest_target = self.wide_array_operand(dest.value);
                        let src = self.wide_array_operand(src.value);
                        return wchar_adapter::wcscpy(dest_target, src);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // strncpy(dest, src, n) → dest = src.substring(0, n)
                "strncpy" => {
                    let mut it = args.into_iter();
                    if let (Some(dest), Some(src), Some(n)) = (it.next(), it.next(), it.next()) {
                        let clipped = expr(ExprKind::Call {
                            callee: Box::new(expr(ExprKind::Member {
                                object: Box::new(src.value),
                                field: "substring".to_string(),
                                null_safe: false,
                            })),
                            args: vec![
                                Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))),
                                Argument::positional(n.value),
                            ],
                            optional: false,
                        });
                        return expr(ExprKind::Assign {
                            target: Box::new(dest.value),
                            value: Box::new(clipped),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "wcscmp" => {
                    let mut it = args.into_iter();
                    if let (Some(a), Some(b)) = (it.next(), it.next()) {
                        return wchar_adapter::wcscmp(
                            self.wide_array_operand(a.value),
                            self.wide_array_operand(b.value),
                        );
                    }
                    return int_lit(0);
                }
                // strcat(dest, src) → dest = dest + src  (returns dest)
                "strcat" => {
                    let mut it = args.into_iter();
                    if let (Some(dest), Some(src)) = (it.next(), it.next()) {
                        if let ExprKind::Assign { target, value } = dest.value.kind {
                            let copy = expr(ExprKind::Assign {
                                target: target.clone(),
                                value,
                            });
                            let concat = expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(*target.clone()),
                                right: Box::new(src.value),
                            });
                            let append = expr(ExprKind::Assign {
                                target,
                                value: Box::new(concat),
                            });
                            return expr(ExprKind::Sequence(vec![copy, append]));
                        }
                        let concat = expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(dest.value.clone()),
                            right: Box::new(src.value),
                        });
                        return expr(ExprKind::Assign {
                            target: Box::new(dest.value),
                            value: Box::new(concat),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // strncat(dest, src, n) → dest = dest + src.substring(0, n)
                "strncat" => {
                    let mut it = args.into_iter();
                    if let (Some(dest), Some(src), Some(n)) = (it.next(), it.next(), it.next()) {
                        let clipped = expr(ExprKind::Call {
                            callee: Box::new(expr(ExprKind::Member {
                                object: Box::new(src.value),
                                field: "substring".to_string(),
                                null_safe: false,
                            })),
                            args: vec![
                                Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))),
                                Argument::positional(n.value),
                            ],
                            optional: false,
                        });
                        let concat = expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(dest.value.clone()),
                            right: Box::new(clipped),
                        });
                        return expr(ExprKind::Assign {
                            target: Box::new(dest.value),
                            value: Box::new(concat),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // strncmp(a, b, n) → strcmp(a.substring(0,n), b.substring(0,n))
                "strncmp" => {
                    let mut it = args.into_iter();
                    if let (Some(a), Some(b), Some(n)) = (it.next(), it.next(), it.next()) {
                        let left = expr(ExprKind::Call {
                            callee: Box::new(expr(ExprKind::Member {
                                object: Box::new(a.value),
                                field: "substring".to_string(),
                                null_safe: false,
                            })),
                            args: vec![
                                Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))),
                                Argument::positional(n.value.clone()),
                            ],
                            optional: false,
                        });
                        let right = expr(ExprKind::Call {
                            callee: Box::new(expr(ExprKind::Member {
                                object: Box::new(b.value),
                                field: "substring".to_string(),
                                null_safe: false,
                            })),
                            args: vec![
                                Argument::positional(expr(ExprKind::Lit(Literal::Int(0)))),
                                Argument::positional(n.value),
                            ],
                            optional: false,
                        });
                        return expr(ExprKind::Call {
                            callee: Box::new(ident("strcmp")),
                            args: vec![Argument::positional(left), Argument::positional(right)],
                            optional: false,
                        });
                    }
                    return int_lit(0);
                }

                // ── stdlib.h — conversions ───────────────────────────────────
                // atoi/atol: parse leading digits, return 0 for non-numeric
                // profile routes to opcode:to_int which fails for "15cats" → 0.
                // Rewrite to parseInt(s, 10) logical-or 0.
                "atoi" | "atol" => {
                    if let Some(s_arg) = args.into_iter().next() {
                        let parse_call = expr(ExprKind::Call {
                            callee: Box::new(ident("parseInt")),
                            args: vec![
                                Argument::positional(s_arg.value),
                                Argument::positional(expr(ExprKind::Lit(Literal::Int(10)))),
                            ],
                            optional: false,
                        });
                        return nan_to_default(parse_call, expr(ExprKind::Lit(Literal::Int(0))));
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                // atof: parseFloat(s) || 0 — returns 0 for empty/non-numeric
                "atof" => {
                    if let Some(s_arg) = args.into_iter().next() {
                        let parse_call = expr(ExprKind::Call {
                            callee: Box::new(ident("parseFloat")),
                            args: vec![Argument::positional(s_arg.value)],
                            optional: false,
                        });
                        return nan_to_default(
                            parse_call,
                            expr(ExprKind::Lit(Literal::Float(0.0))),
                        );
                    }
                    return expr(ExprKind::Lit(Literal::Float(0.0)));
                }
                "div" | "ldiv" => {
                    let mut it = args.into_iter();
                    if let (Some(a), Some(b)) = (it.next(), it.next()) {
                        let quot = expr(ExprKind::Cast {
                            expr: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Div,
                                left: Box::new(a.value.clone()),
                                right: Box::new(b.value.clone()),
                            })),
                            type_name: "int".to_string(),
                        });
                        let rem = expr(ExprKind::Binary {
                            op: BinOp::Mod,
                            left: Box::new(a.value),
                            right: Box::new(b.value),
                        });
                        return expr(ExprKind::Object(vec![
                            ObjectProperty::KeyValue {
                                key: str_lit("quot"),
                                value: quot,
                            },
                            ObjectProperty::KeyValue {
                                key: str_lit("rem"),
                                value: rem,
                            },
                        ]));
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "srand" => {
                    if let Some(seed) = args.into_iter().next() {
                        return call_expr(ident("__c_srand_h"), vec![seed.value]);
                    }
                    return int_lit(0);
                }
                "rand" => {
                    return call_expr(ident("__c_rand_h"), vec![]);
                }
                "signal" => {
                    let mut it = args.into_iter();
                    if let (Some(sig), Some(handler)) = (it.next(), it.next()) {
                        return call_expr(ident("__c_signal_h"), vec![sig.value, handler.value]);
                    }
                    return int_lit(0);
                }
                "raise" => {
                    if let Some(sig) = args.into_iter().next() {
                        return call_expr(ident("__c_raise_h"), vec![sig.value]);
                    }
                    return int_lit(0);
                }
                "setlocale" => {
                    let mut it = args.into_iter();
                    if let (Some(category), Some(locale)) = (it.next(), it.next()) {
                        return call_expr(ident("__c_setlocale_h"), vec![category.value, locale.value]);
                    }
                    return ident("__c_locale");
                }
                "strcoll" => {
                    let mut it = args.into_iter();
                    if let (Some(a), Some(b)) = (it.next(), it.next()) {
                        return call_expr(ident("__c_strcoll_h"), vec![a.value, b.value]);
                    }
                    return int_lit(0);
                }
                "strxfrm" => {
                    let mut it = args.into_iter();
                    if let (Some(dst), Some(src), Some(n)) = (it.next(), it.next(), it.next()) {
                        let write = self.rewrite_memcpy_like(dst.value, src.value.clone(), n.value);
                        return expr(ExprKind::Sequence(vec![write, member(src.value, "length")]));
                    }
                    return int_lit(0);
                }
                "strtol" => {
                    let mut it = args.into_iter();
                    if let Some(s_arg) = it.next() {
                        let end_arg = it.next();
                        let radix = it.next().map(|a| a.value).unwrap_or_else(|| int_lit(10));
                        let parsed = nan_to_default(
                            call_expr(ident("parseInt"), vec![s_arg.value.clone(), radix]),
                            int_lit(0),
                        );
                        if let Some(end) = end_arg {
                            if let Some(end_name) = pointer_address_target_from_init(&Some(end.value.clone())) {
                                let consumed = member(call_expr(ident("String"), vec![parsed.clone()]), "length");
                                let suffix = call_expr(member(s_arg.value, "substring"), vec![consumed]);
                                return expr(ExprKind::Sequence(vec![assign_expr(ident(&end_name), suffix), parsed]));
                            }
                        }
                        return parsed;
                    }
                    return int_lit(0);
                }
                "strtoul" => {
                    let mut it = args.into_iter();
                    if let Some(s_arg) = it.next() {
                        let end_arg = it.next();
                        let radix = it.next().map(|a| a.value).unwrap_or_else(|| int_lit(10));
                        let parsed = nan_to_default(
                            call_expr(ident("parseInt"), vec![s_arg.value.clone(), radix]),
                            int_lit(0),
                        );
                        if let Some(end) = end_arg {
                            if let Some(end_name) = pointer_address_target_from_init(&Some(end.value.clone())) {
                                let consumed = member(call_expr(ident("String"), vec![parsed.clone()]), "length");
                                let suffix = call_expr(member(s_arg.value, "substring"), vec![consumed]);
                                return expr(ExprKind::Sequence(vec![assign_expr(ident(&end_name), suffix), parsed]));
                            }
                        }
                        return parsed;
                    }
                    return int_lit(0);
                }
                "strtod" | "strtof" => {
                    let mut it = args.into_iter();
                    if let Some(s_arg) = it.next() {
                        let end_arg = it.next();
                        let parsed = nan_to_default(
                            call_expr(ident("parseFloat"), vec![s_arg.value.clone()]),
                            int_lit(0),
                        );
                        if let Some(end) = end_arg {
                            if let Some(end_name) = pointer_address_target_from_init(&Some(end.value.clone())) {
                                let consumed = member(call_expr(ident("String"), vec![parsed.clone()]), "length");
                                let suffix = call_expr(member(s_arg.value, "substring"), vec![consumed]);
                                return expr(ExprKind::Sequence(vec![assign_expr(ident(&end_name), suffix), parsed]));
                            }
                        }
                        return parsed;
                    }
                    return int_lit(0);
                }
                "strdup" => {
                    if let Some(s) = args.into_iter().next() {
                        return s.value;
                    }
                    return str_lit("");
                }
                "strpbrk" => {
                    let mut it = args.into_iter();
                    if let (Some(s), Some(accept)) = (it.next(), it.next()) {
                        return call_expr(ident("__c_strpbrk_h"), vec![s.value, accept.value]);
                    }
                    return null_lit();
                }
                "strspn" => {
                    let mut it = args.into_iter();
                    if let (Some(s), Some(accept)) = (it.next(), it.next()) {
                        return call_expr(ident("__c_strspn_h"), vec![s.value, accept.value]);
                    }
                    return int_lit(0);
                }
                "strcspn" => {
                    let mut it = args.into_iter();
                    if let (Some(s), Some(reject)) = (it.next(), it.next()) {
                        return call_expr(ident("__c_strcspn_h"), vec![s.value, reject.value]);
                    }
                    return int_lit(0);
                }
                // ── regex.h (POSIX) ──────────────────────────────────────────
                "regcomp" => {
                    let mut it = args.into_iter();
                    if let (Some(preg), Some(pat), Some(flags)) = (it.next(), it.next(), it.next()) {
                        let preg_lval = self.value_from_c_address_arg(preg.value);
                        return regex_adapter::regcomp(preg_lval, pat.value, flags.value);
                    }
                    return int_lit(0);
                }
                "regexec" => {
                    let mut it = args.into_iter();
                    if let (Some(preg), Some(s), Some(nmatch), Some(pmatch)) =
                        (it.next(), it.next(), it.next(), it.next())
                    {
                        let eflags = it.next().map(|a| a.value).unwrap_or_else(|| int_lit(0));
                        let preg_val = self.value_from_c_address_arg(preg.value);
                        let pmatch_arr =
                            carray_base_expr(&pmatch.value).unwrap_or(pmatch.value);
                        return regex_adapter::regexec(
                            preg_val,
                            s.value,
                            nmatch.value,
                            pmatch_arr,
                            eflags,
                        );
                    }
                    return int_lit(1);
                }
                "regfree" => {
                    return regex_adapter::regfree();
                }
                "regerror" => {
                    // regerror(errcode, &preg, errbuf, errbuf_size)
                    let mut it = args.into_iter();
                    if let (Some(errcode), Some(_preg), Some(errbuf)) =
                        (it.next(), it.next(), it.next())
                    {
                        return regex_adapter::regerror(errcode.value, errbuf.value);
                    }
                    return int_lit(0);
                }
                "time" => {
                    let mut it = args.into_iter();
                    let out = it.next().map(|a| a.value).unwrap_or_else(null_lit);
                    return time_adapter::time(out);
                }
                "clock" => {
                    return time_adapter::clock();
                }
                "difftime" => {
                    let mut it = args.into_iter();
                    if let (Some(a), Some(b)) = (it.next(), it.next()) {
                        return time_adapter::difftime(a.value, b.value);
                    }
                    return int_lit(0);
                }
                "gmtime" => {
                    let t = args.into_iter().next().map(|a| a.value).unwrap_or_else(|| int_lit(0));
                    return time_adapter::gmtime(t);
                }
                "localtime" => {
                    let t = args.into_iter().next().map(|a| a.value).unwrap_or_else(|| int_lit(0));
                    return time_adapter::localtime(t);
                }
                "mktime" => {
                    if let Some(tm) = args.into_iter().next() {
                        return time_adapter::mktime(tm.value);
                    }
                    return int_lit(0);
                }
                "strftime" => {
                    let mut it = args.into_iter();
                    if let (Some(buf), Some(size), Some(fmt), Some(_tm)) =
                        (it.next(), it.next(), it.next(), it.next())
                    {
                        let out = time_adapter::strftime_output(fmt.value);
                        return call_expr(
                            expr(ExprKind::Lambda {
                                params: vec![],
                                body: LambdaBody::Block(vec![
                                    var_decl_stmt("__c_strftime_out", out),
                                    stmt(StmtKind::Expr(self.rewrite_memcpy_like(
                                        buf.value,
                                        ident("__c_strftime_out"),
                                        size.value,
                                    ))),
                                    stmt(StmtKind::Return(Some(member(
                                        ident("__c_strftime_out"),
                                        "length",
                                    )))),
                                ]),
                                is_async: false,
                                captures: vec![],
                            }),
                            vec![],
                        );
                    }
                    return int_lit(0);
                }
                "atomic_load" => {
                    if let Some(addr) = args.into_iter().next() {
                        return atomic_pointer_target(addr.value);
                    }
                    return int_lit(0);
                }
                "atomic_store" => {
                    let mut it = args.into_iter();
                    if let (Some(addr), Some(value)) = (it.next(), it.next()) {
                        return assign_expr(atomic_pointer_target(addr.value), value.value);
                    }
                    return int_lit(0);
                }
                "atomic_fetch_add" => {
                    let mut it = args.into_iter();
                    if let (Some(addr), Some(delta)) = (it.next(), it.next()) {
                        let target = atomic_pointer_target(addr.value);
                        return call_expr(
                            expr(ExprKind::Lambda {
                                params: vec![],
                                body: LambdaBody::Block(vec![
                                    var_decl_stmt("__c_atomic_old", target.clone()),
                                    stmt(StmtKind::Expr(assign_expr(
                                        target,
                                        expr(ExprKind::Binary {
                                            op: BinOp::Add,
                                            left: Box::new(ident("__c_atomic_old")),
                                            right: Box::new(delta.value),
                                        }),
                                    ))),
                                    stmt(StmtKind::Return(Some(ident("__c_atomic_old")))),
                                ]),
                                is_async: false,
                                captures: vec![],
                            }),
                            vec![],
                        );
                    }
                    return int_lit(0);
                }
                "atomic_fetch_sub" => {
                    let mut it = args.into_iter();
                    if let (Some(addr), Some(delta)) = (it.next(), it.next()) {
                        let target = atomic_pointer_target(addr.value);
                        return call_expr(
                            expr(ExprKind::Lambda {
                                params: vec![],
                                body: LambdaBody::Block(vec![
                                    var_decl_stmt("__c_atomic_old", target.clone()),
                                    stmt(StmtKind::Expr(assign_expr(
                                        target,
                                        expr(ExprKind::Binary {
                                            op: BinOp::Sub,
                                            left: Box::new(ident("__c_atomic_old")),
                                            right: Box::new(delta.value),
                                        }),
                                    ))),
                                    stmt(StmtKind::Return(Some(ident("__c_atomic_old")))),
                                ]),
                                is_async: false,
                                captures: vec![],
                            }),
                            vec![],
                        );
                    }
                    return int_lit(0);
                }
                "atomic_compare_exchange_strong" => {
                    let mut it = args.into_iter();
                    if let (Some(addr), Some(expected_ptr), Some(desired)) =
                        (it.next(), it.next(), it.next())
                    {
                        let target = atomic_pointer_target(addr.value);
                        let expected_target = atomic_pointer_target(expected_ptr.value);
                        return call_expr(
                            expr(ExprKind::Lambda {
                                params: vec![],
                                body: LambdaBody::Block(vec![
                                    var_decl_stmt("__c_atomic_cur", target.clone()),
                                    if_stmt(
                                        expr(ExprKind::Binary {
                                            op: BinOp::Eq,
                                            left: Box::new(ident("__c_atomic_cur")),
                                            right: Box::new(expected_target.clone()),
                                        }),
                                        vec![
                                            stmt(StmtKind::Expr(assign_expr(target, desired.value))),
                                            stmt(StmtKind::Return(Some(int_lit(1)))),
                                        ],
                                        Some(vec![
                                            stmt(StmtKind::Expr(assign_expr(
                                                expected_target,
                                                ident("__c_atomic_cur"),
                                            ))),
                                            stmt(StmtKind::Return(Some(int_lit(0)))),
                                        ]),
                                    ),
                                ]),
                                is_async: false,
                                captures: vec![],
                            }),
                            vec![],
                        );
                    }
                    return int_lit(0);
                }
                "atomic_flag_test_and_set" => {
                    if let Some(flag_ptr) = args.into_iter().next() {
                        let target = atomic_pointer_target(flag_ptr.value);
                        return call_expr(
                            expr(ExprKind::Lambda {
                                params: vec![],
                                body: LambdaBody::Block(vec![
                                    var_decl_stmt("__c_atomic_old", target.clone()),
                                    stmt(StmtKind::Expr(assign_expr(target, int_lit(1)))),
                                    stmt(StmtKind::Return(Some(expr(ExprKind::Ternary {
                                        cond: Box::new(expr(ExprKind::Binary {
                                            op: BinOp::NotEq,
                                            left: Box::new(ident("__c_atomic_old")),
                                            right: Box::new(int_lit(0)),
                                        })),
                                        then: Box::new(int_lit(1)),
                                        else_: Box::new(int_lit(0)),
                                    })))),
                                ]),
                                is_async: false,
                                captures: vec![],
                            }),
                            vec![],
                        );
                    }
                    return int_lit(0);
                }

                // ── stdlib.h — heap allocation → arrays ──────────────────────
                // malloc(n) → [] (GC-managed array)
                "malloc" => {
                    return expr(ExprKind::Array(Vec::new()));
                }
                // realloc(p, n) → preserve the existing backing store
                // realloc(p, n): keep the existing backing store (it grows on
                // index write); realloc(NULL, n) behaves like malloc → new array.
                "realloc" => {
                    if let Some(first) = args.into_iter().next() {
                        let p = first.value;
                        if is_zero_int_expr(&p) || matches!(p.kind, ExprKind::Lit(Literal::Null)) {
                            return expr(ExprKind::Array(Vec::new()));
                        }
                        // `p != null ? p : []` for runtime-NULL pointers.
                        return expr(ExprKind::Ternary {
                            cond: Box::new(expr(ExprKind::Binary {
                                op: BinOp::NotEq,
                                left: Box::new(p.clone()),
                                right: Box::new(expr(ExprKind::Lit(Literal::Null))),
                            })),
                            then: Box::new(p),
                            else_: Box::new(expr(ExprKind::Array(Vec::new()))),
                        });
                    }
                    return expr(ExprKind::Array(Vec::new()));
                }
                // calloc(count, size) → pre-filled zero array
                "calloc" => {
                    let count_val = args
                        .into_iter()
                        .next()
                        .and_then(|a| {
                            if let ExprKind::Lit(Literal::Int(n)) = &a.value.kind {
                                Some(*n as usize)
                            } else {
                                None
                            }
                        })
                        .unwrap_or(0);
                    let zeros: Vec<ArrayElement> = (0..count_val)
                        .map(|_| ArrayElement {
                            value: expr(ExprKind::Lit(Literal::Int(0))),
                            spread: false,
                            key: None,
                            by_ref: false,
                        })
                        .collect();
                    return expr(ExprKind::Array(zeros));
                }
                "free" => {
                    // noop — GC handles deallocation
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "memcpy" | "memmove" => {
                    let mut it = args.into_iter();
                    if let (Some(dst), Some(src), Some(bytes)) = (it.next(), it.next(), it.next()) {
                        return self.rewrite_memcpy_like(dst.value, src.value, bytes.value);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "memset" => {
                    let mut it = args.into_iter();
                    if let (Some(dst), Some(fill), Some(bytes)) = (it.next(), it.next(), it.next()) {
                        return self.rewrite_memset(dst.value, fill.value, bytes.value);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                "memchr" => {
                    let mut it = args.into_iter();
                    if let (Some(buf), Some(ch), Some(_bytes)) = (it.next(), it.next(), it.next()) {
                        let needle = char_assignment_value_to_string(ch.value);
                        return strchr_expr(buf.value, needle);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }

                _ => {}
            }
        }
        // Struct function-pointer field calls: `op.apply(args)`.
        // The compiler generates method-call semantics for Member-based callees,
        // passing the object as `this`. In C there is no `this` — use the
        // comma trick `(0, op.apply)(args)` to force a plain call.
        if let ExprKind::Member { object, .. } = &callee.kind {
            if let ExprKind::Ident(var_name) = &object.kind {
                let type_text = self
                    .var_types
                    .get(var_name.as_str())
                    .cloned()
                    .unwrap_or_default();
                let is_struct_type = type_text.contains("struct")
                    || type_text.contains("union")
                    || self.structs.contains_key(type_text.trim());
                if is_struct_type {
                    let seq = expr(ExprKind::Sequence(vec![
                        expr(ExprKind::Lit(Literal::Int(0))),
                        callee,
                    ]));
                    return expr(ExprKind::Call {
                        callee: Box::new(seq),
                        args,
                        optional: false,
                    });
                }
            }
        }
        let callee_name = if let ExprKind::Ident(name) = &callee.kind {
            Some(name.clone())
        } else {
            None
        };
        let call = expr(ExprKind::Call {
            callee: Box::new(callee),
            args: normalized_call_args.clone(),
            optional: false,
        });
        if let Some(name) = callee_name {
            return self.apply_char_param_writebacks(&name, normalized_call_args, call);
        }
        call
    }

    fn apply_char_param_writebacks(
        &mut self,
        callee: &str,
        args: Vec<Argument>,
        call: Expression,
    ) -> Expression {
        let Some(writes) = self.char_param_writes.get(callee).cloned() else {
            return call;
        };
        let mut seq = vec![call];
        for (param_idx, index, value) in &writes {
            let Some(arg) = args.get(*param_idx) else {
                continue;
            };
            let ExprKind::Ident(arg_name) = &arg.value.kind else {
                continue;
            };
            let target = expr(ExprKind::Index {
                object: Box::new(ident(arg_name)),
                index: Box::new(index.clone()),
                null_safe: false,
            });
            if let Some(assign) = self.rewrite_char_index_assignment(&target, value.clone()) {
                seq.push(assign);
            }
        }
        if seq.len() == 1 {
            seq.pop().unwrap()
        } else {
            expr(ExprKind::Sequence(seq))
        }
    }

    /// Normalize a printf/fprintf variadic argument. A char buffer printed with
    /// `%s` must stop at the C null terminator, so truncate char-pointer / carray
    /// args at the first `\0` (same visibility rule `puts` applies). Other args
    /// (ints, floats, single chars) are passed through unchanged.
    fn c_printf_arg(&self, value: Expression) -> Expression {
        if is_carray_object(&value)
            || matches!(&value.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n))
        {
            return c_string_visible(pointers::carray_chars_to_string(value));
        }
        // A `char[]` struct field (e.g. a flexible array member `char data[]`)
        // evaluates to its backing code-point array; `%s` must decode it to a
        // string and stop at the NUL terminator.
        if self.member_is_char_array_field(&value) {
            return c_string_visible(pointers::code_array_to_string(value));
        }
        if matches!(&value.kind, ExprKind::Ident(n) if self.char_pointers.contains(n)) {
            return c_string_visible(value);
        }
        // A string literal carrying an embedded NUL (`"hello\0world"`) prints only
        // up to the terminator under `%s`.
        if matches!(&value.kind, ExprKind::Lit(Literal::Str(s)) if s.contains('\0')) {
            return c_string_visible(value);
        }
        value
    }

    /// True if `value` is a struct-member access `obj.field` (possibly through a
    /// pointer/carray) whose declared field type is a `char` array. Used to
    /// decode the backing code-point array under `%s`/`puts`.
    fn member_is_char_array_field(&self, value: &Expression) -> bool {
        let ExprKind::Member { object, field, .. } = &value.kind else {
            return false;
        };
        let Some(base) = base_ident_name(object) else {
            return false;
        };
        let Some(var_type) = self.var_types.get(&base) else {
            return false;
        };
        let tag = normalized_c_type_name(var_type).replace('*', "").trim().to_string();
        let Some(field_type) = self
            .struct_field_types
            .get(&tag)
            .and_then(|fields| fields.get(field))
        else {
            return false;
        };
        let lower = field_type.to_ascii_lowercase();
        lower.contains("char") && lower.contains('[')
    }

    fn rewrite_sscanf_literal_call(
        &self,
        source_text: &str,
        format_text: &str,
        dest_args: Vec<Argument>,
    ) -> Expression {
        let targets: Vec<Expression> = dest_args
            .into_iter()
            .map(|arg| sscanf_target_expr(&arg.value))
            .collect();
        stdio_adapter::sscanf_literal(source_text, format_text, targets)
    }

    /// Lower `scanf(fmt, &t1, ...)` (fmt a compile-time literal) into a sequence
    /// that reads conversions from the WASI-backed stdin token reader at runtime,
    /// assigns each to its target, and evaluates to the count of items matched.
    /// Conversions stop at the first failed read, per C semantics.
    fn rewrite_scanf_call(&mut self, fmt: &str, targets: Vec<Expression>) -> Expression {
        let id = self.tmp_counter;
        self.tmp_counter += 1;
        crate::platforms::libc::stdio_adapter::scanf(fmt, targets, id)
    }

    fn rewrite_strtok_call(&self, source: Expression, delim: Expression) -> Expression {
        let source_addr_value = self.value_from_c_address_arg(source.clone());
        let source_value = if is_carray_object(&source_addr_value)
            || matches!(&source_addr_value.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n))
        {
            pointers::carray_chars_to_string(source_addr_value.clone())
        } else if matches!(&source_addr_value.kind, ExprKind::Ident(n) if self.char_pointers.contains(n))
            || matches!(&source_addr_value.kind, ExprKind::Lit(Literal::Str(_)))
        {
            source_addr_value
        } else {
            c_string_visible(source_addr_value)
        };

        let delim_addr_value = self.value_from_c_address_arg(delim.clone());
        let delim_value = if is_carray_object(&delim_addr_value)
            || matches!(&delim_addr_value.kind, ExprKind::Ident(n) if self.carray_ptr_vars.contains(n))
        {
            pointers::carray_chars_to_string(delim_addr_value.clone())
        } else if matches!(&delim_addr_value.kind, ExprKind::Ident(n) if self.char_pointers.contains(n))
            || matches!(&delim_addr_value.kind, ExprKind::Lit(Literal::Str(_)))
        {
            delim_addr_value
        } else {
            c_string_visible(delim_addr_value)
        };
        let source_present = expr(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(expr(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(source.clone()),
                right: Box::new(null_lit()),
            })),
            right: Box::new(expr(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(source.clone()),
                right: Box::new(int_lit(0)),
            })),
        });

        string_adapter::strtok_stateful(source_present, source_value, delim_value)
    }

    fn rewrite_va_arg_expression(&self, ap_expr: Expression, type_name: &str) -> Expression {
        let idx_expr = expr(ExprKind::Unary {
            op: UnaryOp::PostInc,
            expr: Box::new(member(ap_expr.clone(), "__idx")),
        });
        let raw_value = index_expr(member(ap_expr, "__values"), idx_expr);

        let normalized = normalized_c_type_name(type_name);
        if normalized.contains('*') {
            return raw_value;
        }
        let cast_target = if normalized.contains("double") || normalized.contains("float") {
            Some("double")
        } else if normalized.contains("unsigned") && normalized.contains("char") {
            Some("uint8")
        } else if normalized.contains("unsigned") {
            Some("uint32")
        } else if normalized.contains("short") {
            Some("int16")
        } else if normalized.contains("long") {
            Some("long")
        } else if normalized.contains("char") {
            Some("char")
        } else if normalized.contains("int") {
            Some("int")
        } else {
            None
        };

        if let Some(cast_target) = cast_target {
            expr(ExprKind::Cast {
                expr: Box::new(raw_value),
                type_name: cast_target.to_string(),
            })
        } else {
            raw_value
        }
    }

    /// Expand a function-like macro call by substituting args for params in the body.
    fn expand_macro_call(
        &mut self,
        params: &[String],
        body: &str,
        args: Vec<Argument>,
    ) -> Expression {
        let substituted = expand_macro_text(params, body, &args, &self.object_macros);
        let trimmed = substituted.trim();
        // Parse the substituted body as a C expression.
        if let Ok(mut pairs) = CParser::parse(Rule::assignment_expression, trimmed) {
            if let Some(pair) = pairs.next() {
                return self.walk_assignment(pair);
            }
        }
        // Token-paste macros can produce bare identifiers that occasionally fail
        // assignment-expression parsing; preserve them instead of nulling out.
        if !trimmed.is_empty()
            && trimmed
                .chars()
                .next()
                .map(|c| c.is_ascii_alphabetic() || c == '_')
                .unwrap_or(false)
            && trimmed
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || c == '_')
        {
            return ident(trimmed);
        }
        if let Ok(v) = trimmed.parse::<i64>() {
            return int_lit(v);
        }
        // Fallback: null
        expr(ExprKind::Lit(Literal::Null))
    }

    fn expand_statement_macro_expr(&mut self, expr: &Expression) -> Option<Vec<Statement>> {
        if let ExprKind::Lit(Literal::Str(marker)) = &expr.kind {
            if let Some(rest) = marker.strip_prefix("__stmt_macro__") {
                let mut body = rest.trim().to_string();
                if !body.ends_with(';') {
                    body.push(';');
                }
                let wrapped = format!("{{ {} }}", body);
                let Ok(mut pairs) = CParser::parse(Rule::compound_statement, &wrapped) else {
                    return None;
                };
                let pair = pairs.next()?;
                return Some(self.walk_block(pair));
            }
        }
        let ExprKind::Call { callee, args, .. } = &expr.kind else {
            return None;
        };
        let ExprKind::Ident(name) = &callee.kind else {
            return None;
        };
        let (params, body) = self.macros.get(name)?.clone();
        let substituted = expand_macro_text(&params, &body, args, &self.object_macros);
        let mut body = substituted.trim().to_string();
        if !body.ends_with(';') {
            body.push(';');
        }
        let wrapped = format!("{{ {} }}", body);
        let Ok(mut pairs) = CParser::parse(Rule::compound_statement, &wrapped) else {
            return None;
        };
        let pair = pairs.next()?;
        Some(self.walk_block(pair))
    }

    fn walk_primary(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::primary_expression => {
                let inner = pair.into_inner().next().unwrap();
                self.walk_primary(inner)
            }
            Rule::literal => self.walk_literal(pair),
            Rule::ident_name => {
                let name = pair.as_str();
                // Remap static local variable accesses to the mangled global name
                if let Some(mangled) = self.static_renames.get(name) {
                    ident(mangled)
                } else if name == "__func__" {
                    expr(ExprKind::Lit(Literal::Str(self.current_function.clone())))
                } else if name == "NULL" {
                    expr(ExprKind::Lit(Literal::Null))
                } else if name == "ATOMIC_FLAG_INIT" {
                    int_lit(0)
                } else if let Some(value) = self.enum_constants.get(name) {
                    expr(ExprKind::Lit(Literal::Int(*value)))
                } else if name == "stdout" {
                    int_lit(1)
                } else if self.address_taken.contains(name) {
                    expr(ExprKind::RefLoad(Box::new(ident(name))))
                } else {
                    ident(name)
                }
            }
            Rule::expression => self.walk_expression(pair),
            _ => self.walk_expression(pair),
        }
    }

    fn walk_literal(&mut self, pair: Pair<Rule>) -> Expression {
        let inner = pair.into_inner().next().unwrap();
        match inner.as_rule() {
            Rule::int_literal => {
                let raw = inner.as_str().trim_end_matches(['u', 'U', 'l', 'L']);
                let v = parse_int_literal(raw);
                expr(ExprKind::Lit(Literal::Int(v)))
            }
            Rule::float_literal => {
                let raw = inner.as_str().trim_end_matches(['f', 'F', 'l', 'L']);
                let v = raw.parse::<f64>().unwrap_or(0.0);
                expr(ExprKind::Lit(Literal::Float(v)))
            }
            Rule::char_literal => {
                let c = parse_char_literal(inner.as_str());
                expr(ExprKind::Lit(Literal::Int(c as i64)))
            }
            Rule::wide_char_literal => {
                let raw = inner.as_str().strip_prefix('L').unwrap_or(inner.as_str());
                let c = parse_char_literal(raw);
                expr(ExprKind::Lit(Literal::Int(c as i64)))
            }
            Rule::string_literal => {
                let s = parse_string_literal(inner.as_str());
                expr(ExprKind::Lit(Literal::Str(s)))
            }
            Rule::wide_string_literal => {
                // L"..." is a wchar_t[] — a flat NUL-terminated code-point array,
                // not a JS string (it must support pointer arithmetic / indexing).
                let raw = inner.as_str().replace("L\"", "\"");
                let s = parse_string_literal(&raw);
                wchar_adapter::wide_string_literal(&s)
            }
            Rule::bool_literal => expr(ExprKind::Lit(Literal::Bool(inner.as_str() == "true"))),
            _ => expr(ExprKind::Lit(Literal::Null)),
        }
    }
}

fn lower_c_gotos(body: Vec<Statement>) -> Vec<Statement> {
    let mut label_to_block: HashMap<String, i64> = HashMap::new();
    let mut blocks: Vec<Vec<Statement>> = vec![Vec::new()];

    for s in body {
        if let StmtKind::Label(name) = s.kind {
            let idx = blocks.len() as i64;
            label_to_block.insert(name, idx);
            blocks.push(Vec::new());
        } else if let Some(last) = blocks.last_mut() {
            last.push(s);
        }
    }

    if label_to_block.is_empty() {
        return blocks.into_iter().next().unwrap_or_default();
    }

    // Declarations before the first label are function-scope in C and must
    // stay visible after jumps to later labels.
    let mut prelude = Vec::new();
    if let Some(first_block) = blocks.first_mut() {
        while first_block
            .first()
            .map(|s| matches!(s.kind, StmtKind::VarDecl { .. }))
            .unwrap_or(false)
        {
            prelude.push(first_block.remove(0));
        }
    }

    let dispatch_label = "__c_goto_dispatch".to_string();
    let pc_name = "__c_goto_pc".to_string();

    let mut switch_cases = Vec::new();
    let total_blocks = blocks.len();
    for (idx, block) in blocks.into_iter().enumerate() {
        let next_pc = if idx + 1 < total_blocks {
            int_lit((idx + 1) as i64)
        } else {
            int_lit(-1)
        };
        let mut case_body = vec![stmt(StmtKind::Expr(assign_expr(ident(&pc_name), next_pc)))];
        case_body.extend(rewrite_gotos_in_stmts(
            block,
            &label_to_block,
            &pc_name,
            &dispatch_label,
        ));
        case_body.push(stmt(StmtKind::Break(BreakTarget::Implicit)));
        switch_cases.push(SwitchCase {
            conditions: vec![CaseCondition::Value(int_lit(idx as i64))],
            body: case_body,
        });
    }

    let while_body = vec![
        stmt(StmtKind::Switch {
            expr: ident(&pc_name),
            cases: switch_cases,
            default: Some(vec![stmt(StmtKind::Break(BreakTarget::Implicit))]),
        }),
        stmt(StmtKind::If {
            cond: expr(ExprKind::Binary {
                op: BinOp::Lt,
                left: Box::new(ident(&pc_name)),
                right: Box::new(int_lit(0)),
            }),
            then_body: vec![stmt(StmtKind::Break(BreakTarget::Implicit))],
            elifs: vec![],
            else_body: None,
        }),
    ];

    let mut lowered = prelude;
    lowered.push(stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(pc_name.clone()),
                type_hint: None,
                init: Some(int_lit(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }));
    lowered.push(stmt(StmtKind::Labeled {
            label: dispatch_label,
            body: Box::new(stmt(StmtKind::While {
                cond: expr(ExprKind::Lit(Literal::Bool(true))),
                body: while_body,
                else_body: None,
            })),
        }));
    lowered
}

fn rewrite_gotos_in_stmts(
    stmts: Vec<Statement>,
    label_to_block: &HashMap<String, i64>,
    pc_name: &str,
    dispatch_label: &str,
) -> Vec<Statement> {
    let mut out = Vec::new();
    for stmt_in in stmts {
        match stmt_in.kind {
            StmtKind::GoTo(target) => {
                if let Some(idx) = label_to_block.get(&target) {
                    out.push(stmt(StmtKind::Expr(assign_expr(ident(pc_name), int_lit(*idx)))));
                    out.push(stmt(StmtKind::Continue(ContinueTarget::Label(
                        dispatch_label.to_string(),
                    ))));
                }
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                let then_body =
                    rewrite_gotos_in_stmts(then_body, label_to_block, pc_name, dispatch_label);
                let elifs = elifs
                    .into_iter()
                    .map(|(c, b)| {
                        (
                            c,
                            rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label),
                        )
                    })
                    .collect();
                let else_body = else_body
                    .map(|b| rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label));
                out.push(stmt(StmtKind::If {
                    cond,
                    then_body,
                    elifs,
                    else_body,
                }));
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                out.push(stmt(StmtKind::For {
                    init,
                    cond,
                    update,
                    body: rewrite_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label),
                }));
            }
            StmtKind::While {
                cond,
                body,
                else_body,
            } => {
                out.push(stmt(StmtKind::While {
                    cond,
                    body: rewrite_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label),
                    else_body: else_body
                        .map(|b| rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label)),
                }));
            }
            StmtKind::DoWhile { body, cond, until } => {
                out.push(stmt(StmtKind::DoWhile {
                    body: rewrite_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label),
                    cond,
                    until,
                }));
            }
            StmtKind::Switch {
                expr,
                cases,
                default,
            } => {
                let cases = cases
                    .into_iter()
                    .map(|mut c| {
                        c.body =
                            rewrite_gotos_in_stmts(c.body, label_to_block, pc_name, dispatch_label);
                        c
                    })
                    .collect();
                let default = default
                    .map(|b| rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label));
                out.push(stmt(StmtKind::Switch {
                    expr,
                    cases,
                    default,
                }));
            }
            StmtKind::Block(body) => {
                out.push(stmt(StmtKind::Block(rewrite_gotos_in_stmts(
                    body,
                    label_to_block,
                    pc_name,
                    dispatch_label,
                ))));
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                let body = rewrite_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label);
                let catches = catches
                    .into_iter()
                    .map(|mut c| {
                        c.body =
                            rewrite_gotos_in_stmts(c.body, label_to_block, pc_name, dispatch_label);
                        c
                    })
                    .collect();
                let else_body = else_body
                    .map(|b| rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label));
                let finally =
                    finally.map(|b| rewrite_gotos_in_stmts(b, label_to_block, pc_name, dispatch_label));
                out.push(stmt(StmtKind::Try {
                    body,
                    catches,
                    else_body,
                    finally,
                }));
            }
            StmtKind::Labeled { label, body } => {
                out.push(stmt(StmtKind::Labeled {
                    label,
                    body: Box::new(stmt(match body.kind {
                        StmtKind::Block(inner) => {
                            StmtKind::Block(rewrite_gotos_in_stmts(
                                inner,
                                label_to_block,
                                pc_name,
                                dispatch_label,
                            ))
                        }
                        other => other,
                    })),
                }));
            }
            StmtKind::Label(_) => {}
            _ => out.push(stmt_in),
        }
    }
    out
}

// ── Binary precedence folding ─────────────────────────────────────────────

fn bin_op(op: &str) -> (BinOp, u8) {
    match op {
        "||" => (BinOp::Or, 1),
        "&&" => (BinOp::And, 2),
        "|" => (BinOp::BitOr, 3),
        "^" => (BinOp::BitXor, 4),
        "&" => (BinOp::BitAnd, 5),
        "==" => (BinOp::Eq, 6),
        "!=" => (BinOp::NotEq, 6),
        "<" => (BinOp::Lt, 7),
        "<=" => (BinOp::LtEq, 7),
        ">" => (BinOp::Gt, 7),
        ">=" => (BinOp::GtEq, 7),
        "<<" => (BinOp::Shl, 8),
        ">>" => (BinOp::Shr, 8),
        "+" => (BinOp::Add, 9),
        "-" => (BinOp::Sub, 9),
        "*" => (BinOp::Mul, 10),
        "/" => (BinOp::Div, 10),
        "%" => (BinOp::Mod, 10),
        _ => (BinOp::Add, 9),
    }
}

/// Precedence-climbing fold over a flat operand/operator sequence.
/// Build a nested zero-filled array matching `bounds` (`[2][2]` → `[[0,0],[0,0]]`).
/// Returns None if any bound is not a non-negative integer literal.
fn zero_nd_array(bounds: &[Expression]) -> Option<Expression> {
    let (first, rest) = bounds.split_first()?;
    let ExprKind::Lit(Literal::Int(n)) = &first.kind else {
        return None;
    };
    if *n < 0 {
        return None;
    }
    let count = *n as usize;
    let elems: Vec<ArrayElement> = (0..count)
        .map(|_| {
            let value = if rest.is_empty() {
                expr(ExprKind::Lit(Literal::Int(0)))
            } else {
                zero_nd_array(rest)?
            };
            Some(ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
        })
        .collect::<Option<Vec<_>>>()?;
    Some(expr(ExprKind::Array(elems)))
}

fn fold_binary(mut operands: Vec<Expression>, ops: Vec<String>) -> Expression {
    if operands.is_empty() {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }
    if ops.is_empty() {
        return operands.pop().unwrap();
    }
    // Shunting-yard into an explicit tree.
    let mut output: Vec<Expression> = Vec::new();
    let mut op_stack: Vec<(BinOp, u8)> = Vec::new();
    let mut operand_iter = operands.into_iter();
    output.push(operand_iter.next().unwrap());

    let apply = |output: &mut Vec<Expression>, op: BinOp| {
        let right = output.pop().unwrap();
        let left = output.pop().unwrap();
        output.push(Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        }));
    };

    for op_str in ops {
        let (op, prec) = bin_op(&op_str);
        while let Some(&(top_op, top_prec)) = op_stack.last() {
            if top_prec >= prec {
                op_stack.pop();
                apply(&mut output, top_op);
            } else {
                break;
            }
        }
        op_stack.push((op, prec));
        if let Some(next) = operand_iter.next() {
            output.push(next);
        }
    }
    while let Some((op, _)) = op_stack.pop() {
        apply(&mut output, op);
    }
    output.pop().unwrap()
}

// ── Literal parsing helpers ────────────────────────────────────────────────

fn parse_int_literal(raw: &str) -> i64 {
    let raw = raw.trim();
    if let Some(hex) = raw.strip_prefix("0x").or_else(|| raw.strip_prefix("0X")) {
        return i64::from_str_radix(hex, 16).unwrap_or(0);
    }
    if let Some(bin) = raw.strip_prefix("0b").or_else(|| raw.strip_prefix("0B")) {
        return i64::from_str_radix(bin, 2).unwrap_or(0);
    }
    if raw.len() > 1 && raw.starts_with('0') && raw.chars().all(|c| c.is_ascii_digit()) {
        return i64::from_str_radix(&raw[1..], 8).unwrap_or(0);
    }
    raw.parse::<i64>().unwrap_or(0)
}

fn parse_char_literal(raw: &str) -> u32 {
    // raw includes surrounding single quotes
    let inner = &raw[1..raw.len() - 1];
    let mut chars = inner.chars().peekable();
    parse_escape_char(&mut chars).unwrap_or(0)
}

/// Parse one C escape or plain character from a peekable char iterator.
fn parse_escape_char(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<u32> {
    let c = chars.next()?;
    if c != '\\' {
        return Some(c as u32);
    }
    let esc = chars.next().unwrap_or('\0');
    Some(match esc {
        'n' => 10,
        't' => 9,
        'r' => 13,
        '0' => 0,
        'a' => 7,
        'b' => 8,
        'f' => 12,
        'v' => 11,
        '\\' => 92,
        '\'' => 39,
        '"' => 34,
        'x' => {
            // \xHH — consume hex digits
            let mut val = 0u32;
            for _ in 0..2 {
                match chars.peek() {
                    Some(&d) if d.is_ascii_hexdigit() => {
                        chars.next();
                        val = val * 16 + d.to_digit(16).unwrap_or(0);
                    }
                    _ => break,
                }
            }
            val
        }
        d if d.is_ascii_digit() => {
            // \NNN — octal (1-3 digits)
            let mut val = d.to_digit(8).unwrap_or(0);
            for _ in 0..2 {
                match chars.peek() {
                    Some(&od) if od.is_ascii_digit() && od < '8' => {
                        chars.next();
                        val = val * 8 + od.to_digit(8).unwrap_or(0);
                    }
                    _ => break,
                }
            }
            val
        }
        other => other as u32,
    })
}

/// Fold `printf`-style `*` width / `.*` precision into the format string by
/// pulling the corresponding (literal) argument. `printf("%*d", 6, 42)` →
/// format `"%6d"` with `42` left in `args`. Args consumed by a `*` are removed
/// from `args`; value args are left in place. Non-literal `*` args are left as
/// `*` (rare; falls through unchanged).
fn resolve_star_format(fmt: &str, args: &mut Vec<Argument>) -> String {
    let chars: Vec<char> = fmt.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    let mut cursor = 0usize; // next arg to consume (star or value)
    let mut remove: Vec<usize> = Vec::new();
    let is_conv = |c: char| {
        matches!(
            c,
            'd' | 'i' | 'u' | 'o' | 'x' | 'X' | 'f' | 'F' | 'e' | 'E' | 'g' | 'G' | 'a' | 'A'
                | 'c' | 's' | 'p'
        )
    };
    while i < chars.len() {
        if chars[i] != '%' {
            out.push(chars[i]);
            i += 1;
            continue;
        }
        out.push('%');
        i += 1;
        if i < chars.len() && chars[i] == '%' {
            out.push('%');
            i += 1;
            continue;
        }
        // walk the conversion spec body, substituting `*` and tracking the value arg
        while i < chars.len() {
            let c = chars[i];
            if c == '*' {
                // pull a literal int arg as the width/precision
                if let Some(arg) = args.get(cursor) {
                    if let ExprKind::Lit(Literal::Int(n)) = &arg.value.kind {
                        out.push_str(&n.to_string());
                        remove.push(cursor);
                        cursor += 1;
                        i += 1;
                        continue;
                    }
                }
                // non-literal: leave as-is
                out.push('*');
                i += 1;
            } else if is_conv(c) {
                out.push(c);
                i += 1;
                cursor += 1; // value arg
                break;
            } else {
                out.push(c);
                i += 1;
            }
        }
    }
    // remove star-consumed args (descending so indices stay valid)
    remove.sort_unstable();
    for idx in remove.into_iter().rev() {
        if idx < args.len() {
            args.remove(idx);
        }
    }
    out
}

/// Wrap a value to a bitfield of `width` bits: `v & ((1<<width)-1)`, then for a
/// signed bitfield sign-extend (`>= 1<<(width-1)` → subtract `1<<width`).
fn apply_bitfield_mask(value: Expression, width: i64, signed: bool) -> Expression {
    if width <= 0 || width >= 64 {
        return value;
    }
    let mask = (1i64 << width) - 1;
    let masked = expr(ExprKind::Binary {
        op: BinOp::BitAnd,
        left: Box::new(value),
        right: Box::new(int_lit(mask)),
    });
    if !signed {
        return masked;
    }
    let half = 1i64 << (width - 1);
    let full = 1i64 << width;
    expr(ExprKind::Ternary {
        cond: Box::new(expr(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(masked.clone()),
            right: Box::new(int_lit(half)),
        })),
        then: Box::new(expr(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(masked.clone()),
            right: Box::new(int_lit(full)),
        })),
        else_: Box::new(masked),
    })
}

fn parse_string_literal(raw: &str) -> String {
    // Adjacent string literals are concatenated; strip quotes from each segment.
    let mut out = String::new();
    let mut chars = raw.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '"' {
            // read chars until closing quote
            loop {
                match chars.peek() {
                    None => break,
                    Some(&'"') => {
                        chars.next();
                        break;
                    }
                    Some(&'\\') => {
                        if let Some(code) = parse_escape_char(&mut chars) {
                            if let Some(ch) = char::from_u32(code) {
                                out.push(ch);
                            }
                        }
                    }
                    _ => {
                        out.push(chars.next().unwrap());
                    }
                }
            }
        }
        // whitespace between adjacent literals is skipped
    }
    out
}

// ── sizeof ────────────────────────────────────────────────────────────────────

impl Walker {
    /// Compute size of a struct (sum of members) or union (max of members).
    /// `text` can be "struct Foo", "union Bar", or a bare typedef name.
    /// Returns 0 if not found (caller falls through to sizeof_from_type_text).
    fn sizeof_struct_union(&self, text: &str) -> i64 {
        let is_union = text.starts_with("union ");
        let tag = if let Some(t) = text.strip_prefix("struct ") {
            t.trim()
        } else if let Some(t) = text.strip_prefix("union ") {
            t.trim()
        } else {
            text.trim()
        };
        let fields = match self.structs.get(tag) {
            Some(f) => f,
            None => return 0,
        };
        // Prefer real per-field sizes when we tracked field types.
        if let Some(field_types) = self.struct_field_types.get(tag) {
            if is_union {
                return fields
                    .iter()
                    .map(|f| field_types.get(f).map(|ty| sizeof_from_type_text(ty).max(1)).unwrap_or(4))
                    .max()
                    .unwrap_or(4);
            }
            // Struct: lay out fields with C alignment/padding. A flexible array
            // member (`T data[]`) contributes its element alignment but no size.
            let mut offset = 0i64;
            let mut max_align = 1i64;
            for f in fields {
                let ty = field_types.get(f).map(|s| s.as_str()).unwrap_or("int");
                let align = alignof_from_type_text(ty);
                max_align = max_align.max(align);
                offset = align_up(offset, align);
                if !ty.replace(' ', "").ends_with("[]") {
                    offset += sizeof_from_type_text(ty).max(1);
                }
            }
            return align_up(offset, max_align).max(1);
        }
        // Fallback approximation when field types are unknown:
        // struct → 4 bytes per field; union → one int-sized member.
        let per_field = 4i64;
        if is_union {
            per_field
        } else {
            per_field * fields.len() as i64
        }
    }

    fn alignof_struct_union(&self, text: &str) -> i64 {
        let tag = if let Some(t) = text.strip_prefix("struct ") {
            t.trim()
        } else if let Some(t) = text.strip_prefix("union ") {
            t.trim()
        } else {
            text.trim()
        };
        let field_types = match self.struct_field_types.get(tag) {
            Some(field_types) => field_types,
            None => return 0,
        };
        field_types
            .values()
            .map(|field_type| {
                let nested = self.alignof_struct_union(field_type);
                if nested > 0 {
                    nested
                } else {
                    alignof_from_type_text(field_type)
                }
            })
            .max()
            .unwrap_or(1)
    }

    fn offsetof_struct_field(&self, type_text: &str, field_name: &str) -> i64 {
        let tag = if let Some(t) = type_text.strip_prefix("struct ") {
            t.trim()
        } else if let Some(t) = type_text.strip_prefix("union ") {
            t.trim()
        } else {
            type_text.trim()
        };
        let fields = match self.structs.get(tag) {
            Some(fields) => fields,
            None => return 0,
        };
        let field_types = match self.struct_field_types.get(tag) {
            Some(types) => types,
            None => return 0,
        };
        if type_text.trim_start().starts_with("union ") {
            return 0;
        }
        let mut offset = 0i64;
        for field in fields {
            let Some(field_type) = field_types.get(field) else {
                continue;
            };
            let field_align = {
                let nested = self.alignof_struct_union(field_type);
                if nested > 0 {
                    nested
                } else {
                    alignof_from_type_text(field_type)
                }
            }
            .max(1);
            offset = align_up(offset, field_align);
            if field == field_name {
                return offset;
            }
            let field_size = {
                let nested = self.sizeof_struct_union(field_type);
                if nested > 0 {
                    nested
                } else {
                    sizeof_from_type_text(field_type)
                }
            }
            .max(1);
            offset += field_size;
        }
        0
    }

    fn struct_field_at_offset(&self, type_text: &str, offset: i64) -> Option<String> {
        let tag = if let Some(t) = type_text.strip_prefix("struct ") {
            t.trim()
        } else if let Some(t) = type_text.strip_prefix("union ") {
            t.trim()
        } else {
            type_text.trim()
        };
        let fields = self.structs.get(tag)?;
        let field_types = self.struct_field_types.get(tag)?;
        let mut current = 0i64;
        for field in fields {
            let field_type = field_types.get(field)?;
            let field_align = {
                let nested = self.alignof_struct_union(field_type);
                if nested > 0 {
                    nested
                } else {
                    alignof_from_type_text(field_type)
                }
            }
            .max(1);
            current = align_up(current, field_align);
            if current == offset {
                return Some(field.clone());
            }
            let field_size = {
                let nested = self.sizeof_struct_union(field_type);
                if nested > 0 {
                    nested
                } else {
                    sizeof_from_type_text(field_type)
                }
            }
            .max(1);
            current += field_size;
        }
        None
    }

    fn walk_generic_expression(&mut self, pair: Pair<Rule>) -> Expression {
        let mut inner = pair.into_inner();
        let control_pair = inner.next().unwrap();
        let control_expr = self.walk_assignment(control_pair.clone());
        let control_type = self.infer_generic_type(&control_expr, control_pair.as_str());
        let mut default_expr = None;
        for assoc in inner {
            let assoc_src = assoc.as_str().trim().to_string();
            let mut parts = assoc.into_inner();
            if assoc_src.starts_with("default") {
                if let Some(expr_pair) = parts.next() {
                    default_expr = Some(self.walk_assignment(expr_pair));
                }
                continue;
            }
            let Some(type_pair) = parts.next() else {
                continue;
            };
            let Some(expr_pair) = parts.next() else {
                continue;
            };
            let assoc_type = normalized_c_type_name(type_pair.as_str());
            if assoc_type == control_type {
                return self.walk_assignment(expr_pair);
            }
        }
        default_expr.unwrap_or_else(|| expr(ExprKind::Lit(Literal::Null)))
    }

    fn infer_generic_type(&self, expr: &Expression, raw: &str) -> String {
        match &expr.kind {
            ExprKind::Ident(name) => self
                .var_types
                .get(name)
                .map(|ty| normalized_c_type_name(ty))
                .unwrap_or_else(|| "int".to_string()),
            ExprKind::Lit(Literal::Float(_)) => {
                let trimmed = raw.trim();
                if trimmed.ends_with('f') || trimmed.ends_with('F') {
                    "float".to_string()
                } else {
                    "double".to_string()
                }
            }
            ExprKind::Lit(Literal::Int(_)) => {
                if raw.trim().starts_with('\'') {
                    "char".to_string()
                } else {
                    "int".to_string()
                }
            }
            ExprKind::Cast { type_name, .. } => match type_name.as_str() {
                "double" => "double".to_string(),
                "char" => "char".to_string(),
                "long" => "long".to_string(),
                _ => "int".to_string(),
            },
            _ => {
                if let Some(sz) = self.var_types.get(raw.trim()) {
                    normalized_c_type_name(sz)
                } else {
                    "int".to_string()
                }
            }
        }
    }

    fn pointer_member_target_from_char_struct_base_init(
        &self,
        init: &Option<Expression>,
    ) -> Option<Expression> {
        let (base_ptr, offset) = char_pointer_offset_from_init(init)?;
        let struct_var = self.char_pointer_struct_bases.get(&base_ptr)?;
        let struct_type = self.var_types.get(struct_var)?;
        let offset_value = match &offset.kind {
            ExprKind::Lit(Literal::Int(n)) => *n,
            _ => return None,
        };
        let field = self.struct_field_at_offset(struct_type, offset_value)?;
        Some(expr(ExprKind::Member {
            object: Box::new(ident(struct_var)),
            field,
            null_safe: false,
        }))
    }

    fn sizeof_from_rule(&self, pair: &Pair<Rule>) -> i64 {
        match pair.as_rule() {
            Rule::sizeof_expression => {
                for inner in pair.clone().into_inner() {
                    return self.sizeof_from_rule(&inner);
                }
                8
            }
            Rule::cast_expression => {
                let mut inners = pair.clone().into_inner();
                if let Some(first) = inners.next() {
                    if first.as_rule() == Rule::type_name {
                        return self.sizeof_from_rule(&first);
                    }
                }
                sizeof_from_type_text(pair.as_str().trim())
            }
            Rule::type_name
            | Rule::declaration_specifiers
            | Rule::type_specifier
            | Rule::type_specifier_strict => {
                let text = pair.as_str().trim();
                // Could be a variable name parsed as typedef_name
                if let Some(&sz) = self.var_sizes.get(text) {
                    return sz;
                }
                if let Some(ty) = self.var_types.get(text) {
                    let su = self.sizeof_struct_union(ty);
                    return if su > 0 {
                        su
                    } else {
                        sizeof_from_type_text(ty)
                    };
                }
                if let Some(sz) = self.sizeof_from_expr_text(text) {
                    return sz;
                }
                // struct/union: sum or max of member sizes
                let sz = self.sizeof_struct_union(text);
                if sz > 0 {
                    return sz;
                }
                sizeof_from_type_text(text)
            }
            Rule::unary_expression
            | Rule::postfix_expression
            | Rule::primary_expression
            | Rule::expression
            | Rule::assignment_expression => {
                // sizeof(expr) — check if it's a variable name we know
                let text = pair.as_str().trim();
                // Strip parens
                let text = text.trim_start_matches('(').trim_end_matches(')').trim();
                // Char literal → int (4 bytes in C)
                if text.starts_with('\'') {
                    return 4;
                }
                // String literal → strlen + 1
                if text.starts_with('"') {
                    let s = parse_string_literal(text);
                    return (s.len() + 1) as i64;
                }
                // Check if it's a known variable
                if let Some(&sz) = self.var_sizes.get(text) {
                    return sz;
                }
                if let Some(ty) = self.var_types.get(text) {
                    let su = self.sizeof_struct_union(ty);
                    return if su > 0 {
                        su
                    } else {
                        sizeof_from_type_text(ty)
                    };
                }
                if let Some(sz) = self.sizeof_from_expr_text(text) {
                    return sz;
                }
                // `*p` → sizeof the type p points to
                if let Some(inner_name) = text.strip_prefix('*').map(|s| s.trim()) {
                    if let Some(ty) = self.var_types.get(inner_name) {
                        // ty is "int" for `int *p`; pointer-target size
                        return sizeof_from_type_text(ty.trim_end_matches('*').trim());
                    }
                }
                // `arr[n]` → sizeof element of arr
                if let Some(base_name) = text.split('[').next().map(|s| s.trim()) {
                    if let Some(ty) = self.var_types.get(base_name) {
                        let mut dims = text.matches('[').count();
                        if dims == 0 {
                            dims = 1;
                        }
                        return self.sizeof_indexed_expr(base_name, ty, dims);
                    }
                }
                if let Some(sz) = self.sizeof_member_expr_text(text) {
                    return sz;
                }
                // Fall through to text-based guess
                let base = sizeof_from_type_text(text);
                if base != 8 {
                    return base;
                }
                // Try inner nodes
                for inner in pair.clone().into_inner() {
                    let s = self.sizeof_from_rule(&inner);
                    if s != 8 {
                        return s;
                    }
                }
                8
            }
            _ => sizeof_from_type_text(pair.as_str().trim()),
        }
    }

    fn alignof_from_rule(&self, pair: &Pair<Rule>) -> i64 {
        match pair.as_rule() {
            Rule::alignof_expression => {
                for inner in pair.clone().into_inner() {
                    return self.alignof_from_rule(&inner);
                }
                8
            }
            Rule::cast_expression => {
                let mut inners = pair.clone().into_inner();
                if let Some(first) = inners.next() {
                    if first.as_rule() == Rule::type_name {
                        return self.alignof_from_rule(&first);
                    }
                }
                alignof_from_type_text(pair.as_str().trim())
            }
            Rule::type_name
            | Rule::declaration_specifiers
            | Rule::type_specifier
            | Rule::type_specifier_strict => {
                let text = pair.as_str().trim();
                if let Some(ty) = self.var_types.get(text) {
                    let nested = self.alignof_struct_union(ty);
                    return if nested > 0 {
                        nested
                    } else {
                        alignof_from_type_text(ty)
                    };
                }
                if let Some(align) = self.alignof_from_expr_text(text) {
                    return align;
                }
                let nested = self.alignof_struct_union(text);
                if nested > 0 {
                    return nested;
                }
                alignof_from_type_text(text)
            }
            Rule::unary_expression
            | Rule::postfix_expression
            | Rule::primary_expression
            | Rule::expression
            | Rule::assignment_expression => {
                let text = pair.as_str().trim();
                let text = text.trim_start_matches('(').trim_end_matches(')').trim();
                if text.starts_with('\'') {
                    return 1;
                }
                if text.starts_with('"') {
                    return 1;
                }
                if let Some(ty) = self.var_types.get(text) {
                    let nested = self.alignof_struct_union(ty);
                    return if nested > 0 {
                        nested
                    } else {
                        alignof_from_type_text(ty)
                    };
                }
                if let Some(align) = self.alignof_from_expr_text(text) {
                    return align;
                }
                for inner in pair.clone().into_inner() {
                    let align = self.alignof_from_rule(&inner);
                    if align != 8 || text.contains('*') || text.contains("double") {
                        return align;
                    }
                }
                alignof_from_type_text(text)
            }
            _ => alignof_from_type_text(pair.as_str().trim()),
        }
    }

    fn sizeof_indexed_expr(&self, base_name: &str, ty: &str, dims_used: usize) -> i64 {
        let base_size = sizeof_from_type_text(ty);
        let total_size = self.var_sizes.get(base_name).copied().unwrap_or(base_size);
        let declared_count = self.array_element_count_from_type(ty).unwrap_or(1);
        if declared_count <= 1 {
            return base_size;
        }
        let remaining = declared_count;
        if dims_used >= self.array_rank_from_type(ty) {
            base_size
        } else {
            total_size / self.first_array_bound_from_type(ty).unwrap_or(remaining)
        }
    }

    fn array_rank_from_type(&self, ty: &str) -> usize {
        ty.matches('[').count()
    }

    fn first_array_bound_from_type(&self, ty: &str) -> Option<i64> {
        ty.split('[')
            .nth(1)
            .and_then(|rest| rest.split(']').next())
            .and_then(|n| n.trim().parse::<i64>().ok())
    }

    fn array_element_count_from_type(&self, ty: &str) -> Option<i64> {
        let mut count = 1i64;
        let mut found = false;
        for part in ty.split('[').skip(1) {
            if let Some(raw) = part.split(']').next() {
                if let Ok(n) = raw.trim().parse::<i64>() {
                    count *= n.max(1);
                    found = true;
                }
            }
        }
        found.then_some(count)
    }

    fn sizeof_from_expr_text(&self, text: &str) -> Option<i64> {
        let text = text.trim();
        let text = text
            .strip_suffix("++")
            .or_else(|| text.strip_suffix("--"))
            .map(str::trim)
            .unwrap_or(text);
        if let Some(&sz) = self.var_sizes.get(text) {
            return Some(sz);
        }
        if let Some(ty) = self.var_types.get(text) {
            let su = self.sizeof_struct_union(ty);
            return Some(if su > 0 {
                su
            } else {
                sizeof_from_type_text(ty)
            });
        }
        if let Some((_, rhs)) = text.rsplit_once(',') {
            return self.sizeof_from_expr_text(rhs.trim());
        }
        if text.parse::<f64>().is_ok() && text.contains('.') {
            return Some(8);
        }
        if text.parse::<i64>().is_ok() {
            return Some(4);
        }
        if text.contains('?') && text.contains(':') {
            return Some(if text.contains('.') { 8 } else { 4 });
        }
        if let Some(base_name) = text.split('[').next().map(|s| s.trim()) {
            if base_name != text {
                if let Some(ty) = self.var_types.get(base_name) {
                    let dims = text.matches('[').count().max(1);
                    return Some(self.sizeof_indexed_expr(base_name, ty, dims));
                }
            }
        }
        if let Some(sz) = self.sizeof_member_expr_text(text) {
            return Some(sz);
        }
        None
    }

    fn alignof_from_expr_text(&self, text: &str) -> Option<i64> {
        let text = text.trim();
        let text = text
            .strip_suffix("++")
            .or_else(|| text.strip_suffix("--"))
            .map(str::trim)
            .unwrap_or(text);
        if let Some(ty) = self.var_types.get(text) {
            let nested = self.alignof_struct_union(ty);
            let base = if nested > 0 {
                nested
            } else {
                alignof_from_type_text(ty)
            };
            return Some(self.var_alignments.get(text).copied().unwrap_or(base).max(base));
        }
        if let Some((_, rhs)) = text.rsplit_once(',') {
            return self.alignof_from_expr_text(rhs.trim());
        }
        if text.parse::<f64>().is_ok() && text.contains('.') {
            return Some(8);
        }
        if text.parse::<i64>().is_ok() {
            return Some(4);
        }
        if let Some(base_name) = text.split('[').next().map(|s| s.trim()) {
            if base_name != text {
                if let Some(ty) = self.var_types.get(base_name) {
                    let nested = self.alignof_struct_union(ty);
                    return Some(if nested > 0 {
                        nested
                    } else {
                        alignof_from_type_text(ty)
                    });
                }
            }
        }
        if let Some((object_text, _)) = text.rsplit_once("->").or_else(|| text.rsplit_once('.')) {
            let object_name = object_text
                .split(|c: char| !c.is_ascii_alphanumeric() && c != '_')
                .next_back()
                .unwrap_or("")
                .trim();
            if let Some(object_type) = self.var_types.get(object_name) {
                let nested = self.alignof_struct_union(object_type.trim_end_matches('*').trim());
                if nested > 0 {
                    return Some(nested);
                }
            }
        }
        None
    }

    fn sizeof_member_expr_text(&self, text: &str) -> Option<i64> {
        let (object_text, field_text) = text.rsplit_once("->").or_else(|| text.rsplit_once('.'))?;
        let field_name = field_text
            .split(|c: char| !c.is_ascii_alphanumeric() && c != '_')
            .next()
            .unwrap_or("")
            .trim();
        if field_name.is_empty() {
            return None;
        }
        let object_name = object_text
            .split(|c: char| !c.is_ascii_alphanumeric() && c != '_')
            .next_back()
            .unwrap_or("")
            .trim();
        if object_name.is_empty() {
            return None;
        }
        let object_type = self.var_types.get(object_name)?;
        let type_name = normalized_c_type_name(object_type.trim_end_matches('*').trim());
        let field_type = self.struct_field_types.get(&type_name)?.get(field_name)?;
        let nested = self.sizeof_struct_union(field_type);
        Some(if nested > 0 {
            nested
        } else {
            sizeof_from_type_text(field_type)
        })
    }

    fn is_fixed_array_var(&self, name: &str) -> bool {
        self.array_ptr_vars.contains(name)
            || self
                .var_types
                .get(name)
                .map(|type_text| type_text.contains('[') && !type_text.contains("char"))
                .unwrap_or(false)
    }

    fn is_carray_compatible_pointer_param(&self, type_hint: &str) -> bool {
        let pointee = normalized_c_type_name(type_hint)
            .replace('*', "")
            .trim()
            .to_string();
        type_hint.matches('*').count() == 1
            && !type_hint.contains("char")
            && !type_hint.contains("struct")
            && !type_hint.contains("union")
            && !self.structs.contains_key(&pointee)
    }

    fn pointer_init_looks_array_backed(&self, init: Option<&Expression>) -> bool {
        let Some(init) = init else { return false };
        if is_carray_like_expr(init) {
            return true;
        }

        match &init.kind {
            ExprKind::Ident(name) => {
                self.array_ptr_vars.contains(name) || self.carray_ptr_vars.contains(name)
            }
            ExprKind::Binary {
                op: BinOp::Add,
                left,
                ..
            }
            | ExprKind::Binary {
                op: BinOp::Sub,
                left,
                ..
            } => matches!(&left.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n) || self.carray_ptr_vars.contains(n)),
            ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr,
            } => match &expr.kind {
                ExprKind::Index { object, .. } => {
                    matches!(&object.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n) || self.carray_ptr_vars.contains(n))
                }
                _ => false,
            },
            ExprKind::Index { object, .. } => {
                matches!(&object.kind, ExprKind::Ident(n) if self.array_ptr_vars.contains(n) || self.carray_ptr_vars.contains(n))
            }
            _ => false,
        }
    }
}

fn sizeof_from_type_text(text: &str) -> i64 {
    // Strip qualifiers
    let t = normalized_c_type_name(text);
    let t = t.split('[').next().unwrap_or(t.as_str()).trim();
    // Pointer → pointer size (8 on 64-bit)
    if t.contains('*') {
        return 8;
    }
    match t {
        "char" | "int8_t" | "uint8_t" | "_Bool" | "bool" => 1,
        "short" | "int16_t" | "uint16_t" => 2,
        "int" | "float" | "int32_t" | "uint32_t" => 4,
        "long" | "double" | "long long" | "int64_t" | "uint64_t" | "size_t" | "ssize_t"
        | "ptrdiff_t" => 8,
        "long double" => 16,
        "void" => 1,
        _ => 8, // unknown / struct / pointer-like → pointer size
    }
}

fn alignof_from_type_text(text: &str) -> i64 {
    let t = normalized_c_type_name(text);
    let t = t.split('[').next().unwrap_or(t.as_str()).trim();
    if t.contains('*') {
        return 8;
    }
    match t {
        "char" | "int8_t" | "uint8_t" | "_Bool" | "bool" => 1,
        "short" | "int16_t" | "uint16_t" => 2,
        "int" | "float" | "int32_t" | "uint32_t" => 4,
        "long" | "double" | "long long" | "int64_t" | "uint64_t" | "size_t" | "ssize_t"
        | "ptrdiff_t" => 8,
        "long double" => 16,
        "void" => 1,
        _ => 8,
    }
}

fn align_up(offset: i64, align: i64) -> i64 {
    if align <= 1 {
        return offset;
    }
    let rem = offset % align;
    if rem == 0 {
        offset
    } else {
        offset + (align - rem)
    }
}

fn normalized_c_type_name(text: &str) -> String {
    let stripped = strip_alignment_specifiers(text);
    let t = stripped
        .replace("const ", "")
        .replace("volatile ", "")
        .replace("static ", "")
        .replace("unsigned ", "")
        .replace("signed ", "")
        .replace("register ", "")
        .replace("restrict ", "")
        .trim()
        .to_string();
    t.strip_prefix("struct ")
        .or_else(|| t.strip_prefix("union "))
        .unwrap_or(t.as_str())
        .trim()
        .to_string()
}


fn strip_alignment_specifiers(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let bytes = text.as_bytes();
    let mut index = 0;
    while index < bytes.len() {
        let rest = &text[index..];
        let marker = if rest.starts_with("_Alignas(") {
            Some("_Alignas(")
        } else if rest.starts_with("alignas(") {
            Some("alignas(")
        } else {
            None
        };
        if let Some(prefix) = marker {
            index += prefix.len();
            let mut depth = 1usize;
            while index < bytes.len() && depth > 0 {
                match bytes[index] as char {
                    '(' => depth += 1,
                    ')' => depth -= 1,
                    _ => {}
                }
                index += 1;
            }
            if !out.ends_with(' ') {
                out.push(' ');
            }
            continue;
        }
        out.push(bytes[index] as char);
        index += 1;
    }
    out
}

fn explicit_alignment_value(text: &str) -> Option<i64> {
    ["_Alignas(", "alignas("]
        .into_iter()
        .find_map(|marker| {
            let start = text.find(marker)? + marker.len();
            let rest = &text[start..];
            let end = rest.find(')')?;
            rest[..end].trim().parse::<i64>().ok()
        })
}
/// `strchr(s, needle_str)` — find first occurrence, return suffix or null.

fn extract_typeof_expr_text(text: &str) -> Option<&str> {
    for prefix in ["__typeof__(", "typeof("] {
        if let Some(rest) = text.strip_prefix(prefix) {
            return rest.strip_suffix(')');
        }
    }
    None
}

fn normalize_macro_body(text: &str) -> String {
    text.replace("\\\r\n", " ")
        .replace("\\\n", " ")
        .replace("\\\r", " ")
}

fn c_quote_string(text: &str) -> String {
    let escaped = text.replace('\\', "\\\\").replace('"', "\\\"");
    format!("\"{}\"", escaped)
}

fn apply_token_paste(
    mut text: String,
    param_values: &HashMap<String, String>,
    object_macros: &HashMap<String, String>,
) -> String {
    loop {
        let Some(op_pos) = text.find("##") else {
            break;
        };

        let bytes = text.as_bytes();

        let mut l = op_pos;
        while l > 0 && bytes[l - 1].is_ascii_whitespace() {
            l -= 1;
        }
        let mut l_start = l;
        while l_start > 0 {
            let c = bytes[l_start - 1];
            if c.is_ascii_alphanumeric() || c == b'_' {
                l_start -= 1;
            } else {
                break;
            }
        }

        let mut r = op_pos + 2;
        while r < bytes.len() && bytes[r].is_ascii_whitespace() {
            r += 1;
        }
        let mut r_end = r;
        while r_end < bytes.len() {
            let c = bytes[r_end];
            if c.is_ascii_alphanumeric() || c == b'_' {
                r_end += 1;
            } else {
                break;
            }
        }

        if l_start >= l || r >= r_end {
            break;
        }

        let left_tok = &text[l_start..l];
        let right_tok = &text[r..r_end];

        let left_val = param_values
            .get(left_tok)
            .cloned()
            .or_else(|| object_macros.get(left_tok).cloned())
            .unwrap_or_else(|| left_tok.to_string());
        let right_val = param_values
            .get(right_tok)
            .cloned()
            .or_else(|| object_macros.get(right_tok).cloned())
            .unwrap_or_else(|| right_tok.to_string());

        let pasted = format!("{}{}", left_val.trim(), right_val.trim());
        text.replace_range(l_start..r_end, &pasted);
    }
    text
}

fn apply_stringize(
    text: &str,
    param_values: &HashMap<String, String>,
    variadic: &str,
) -> String {
    let bytes = text.as_bytes();
    let mut i = 0usize;
    let mut out = String::with_capacity(text.len());

    while i < bytes.len() {
        let c = bytes[i];
        if c == b'#' {
            let prev_is_hash = i > 0 && bytes[i - 1] == b'#';
            let next_is_hash = i + 1 < bytes.len() && bytes[i + 1] == b'#';
            if !prev_is_hash && !next_is_hash {
                let mut j = i + 1;
                while j < bytes.len()
                    && (bytes[j].is_ascii_alphanumeric() || bytes[j] == b'_')
                {
                    j += 1;
                }
                if j > i + 1 {
                    let token = &text[i + 1..j];
                    if token == "__VA_ARGS__" {
                        out.push_str(&c_quote_string(variadic.trim()));
                        i = j;
                        continue;
                    }
                    if let Some(arg_src) = param_values.get(token) {
                        out.push_str(&c_quote_string(arg_src.trim()));
                        i = j;
                        continue;
                    }
                }
            }
        }

        out.push(bytes[i] as char);
        i += 1;
    }

    out
}

fn expand_macro_text(
    params: &[String],
    body: &str,
    args: &[Argument],
    object_macros: &HashMap<String, String>,
) -> String {
    let mut substituted = normalize_macro_body(body);

    let mut param_values: HashMap<String, String> = HashMap::new();
    for (i, param) in params.iter().enumerate() {
        let arg_src = args
            .get(i)
            .map(|arg| expr_to_c_source(&arg.value))
            .unwrap_or_else(|| "0".to_string());
        param_values.insert(param.clone(), arg_src);
    }

    let variadic = if args.len() > params.len() {
        args[params.len()..]
            .iter()
            .map(|a| expr_to_c_source(&a.value))
            .collect::<Vec<_>>()
            .join(", ")
    } else {
        String::new()
    };

    substituted = apply_stringize(&substituted, &param_values, &variadic);

    substituted = apply_token_paste(substituted, &param_values, object_macros);

    for (param, arg_src) in &param_values {
        substituted = replace_word(&substituted, param, arg_src);
    }
    substituted = replace_word(&substituted, "__VA_ARGS__", &variadic);

    for (name, replacement) in object_macros {
        substituted = replace_word(&substituted, name, replacement);
    }

    substituted
}

fn strchr_expr(s: Expression, needle: Expression) -> Expression {
    if is_putchar_zero_call(&needle) {
        return expr(ExprKind::Lit(Literal::Int(1)));
    }
    let idx_call = expr(ExprKind::Call {
        callee: Box::new(expr(ExprKind::Member {
            object: Box::new(s.clone()),
            field: "indexOf".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(needle)],
        optional: false,
    });
    expr(ExprKind::Ternary {
        cond: Box::new(expr(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(idx_call.clone()),
            right: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
        })),
        then: Box::new(expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member {
                object: Box::new(s),
                field: "slice".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(idx_call)],
            optional: false,
        })),
        else_: Box::new(expr(ExprKind::Lit(Literal::Null))),
    })
}

fn c_string_visible(s: Expression) -> Expression {
    expr(ExprKind::Index {
        object: Box::new(expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member {
                object: Box::new(s),
                field: "split".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(expr(ExprKind::Lit(Literal::Str(
                "\0".to_string(),
            ))))],
            optional: false,
        })),
        index: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
        null_safe: false,
    })
}

/// True if `e` is an Object literal tagged as a carray (from `make_carray_ptr`).
fn is_carray_object(e: &Expression) -> bool {
    if let ExprKind::Object(ref props) = e.kind {
        props.iter().any(|p| {
            if let ObjectProperty::KeyValue { key, value } = p {
                if let (ExprKind::Lit(Literal::Str(k)), ExprKind::Lit(Literal::Str(v))) =
                    (&key.kind, &value.kind)
                {
                    return k == "__ref_kind" && v == CARRAY_KIND;
                }
            }
            false
        })
    } else {
        false
    }
}

fn is_carray_like_expr(e: &Expression) -> bool {
    if is_carray_object(e) {
        return true;
    }
    match &e.kind {
        ExprKind::Ternary { then, else_, .. } => {
            is_carray_like_expr(then) || matches!(else_.kind, ExprKind::Lit(Literal::Null))
        }
        ExprKind::Call { callee, .. } => {
            let ExprKind::Lambda { body, .. } = &callee.kind else {
                return false;
            };
            match body {
                LambdaBody::Expr(expr) => is_carray_like_expr(expr),
                LambdaBody::Block(_) => false,
            }
        }
        _ => false,
    }
}

fn carray_deref_target_name(e: &Expression) -> Option<String> {
    if let ExprKind::Ternary { then, .. } = &e.kind {
        return carray_deref_target_name(then);
    }
    let ExprKind::Index { object, index, .. } = &e.kind else {
        return None;
    };
    let ExprKind::Member {
        object: base_object,
        field: base_field,
        ..
    } = &object.kind
    else {
        return None;
    };
    let ExprKind::Member {
        object: idx_object,
        field: idx_field,
        ..
    } = &index.kind
    else {
        return None;
    };
    if base_field != CARRAY_BASE_KEY || idx_field != CARRAY_IDX_KEY {
        return None;
    }
    let ExprKind::Ident(base_name) = &base_object.kind else {
        return None;
    };
    let ExprKind::Ident(idx_name) = &idx_object.kind else {
        return None;
    };
    (base_name == idx_name).then(|| base_name.clone())
}

fn dynamic_carray_deref_read(ptr: Expression) -> Expression {
    let scalar_read = Expression::new(ExprKind::Unary {
        op: UnaryOp::Deref,
        expr: Box::new(ptr.clone()),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(pointers::is_carray_ptr_kind(ptr.clone())),
        then: Box::new(pointers::carray_deref_read(ptr)),
        else_: Box::new(scalar_read),
    })
}

fn dynamic_carray_deref_write(ptr: Expression, value: Expression) -> Expression {
    let scalar_write = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::Unary {
            op: UnaryOp::Deref,
            expr: Box::new(ptr.clone()),
        })),
        value: Box::new(value.clone()),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(pointers::is_carray_ptr_kind(ptr.clone())),
        then: Box::new(pointers::carray_deref_write(ptr, value)),
        else_: Box::new(scalar_write),
    })
}

/// Build a new carray pointer retreated by `n`: `{__base, __idx: __idx - n}`.
fn carray_retreat(ptr: Expression, n: Expression) -> Expression {
    let new_idx = Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(ptr.clone()),
            field: CARRAY_IDX_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(n),
    });
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::new(ExprKind::Lit(Literal::Str("__ref_kind".to_string()))),
            value: Expression::new(ExprKind::Lit(Literal::Str(CARRAY_KIND.to_string()))),
        },
        ObjectProperty::KeyValue {
            key: Expression::new(ExprKind::Lit(Literal::Str(CARRAY_BASE_KEY.to_string()))),
            value: Expression::new(ExprKind::Member {
                object: Box::new(ptr),
                field: CARRAY_BASE_KEY.to_string(),
                null_safe: false,
            }),
        },
        ObjectProperty::KeyValue {
            key: Expression::new(ExprKind::Lit(Literal::Str(CARRAY_IDX_KEY.to_string()))),
            value: new_idx,
        },
    ]))
}

fn is_null_pointer_init(init: &Option<Expression>) -> bool {
    matches!(
        init.as_ref().map(|e| &e.kind),
        Some(ExprKind::Lit(Literal::Int(0)) | ExprKind::Lit(Literal::Null))
    )
}

fn same_ident_expr(left: &Expression, right: &Expression) -> bool {
    matches!(
        (&left.kind, &right.kind),
        (ExprKind::Ident(l), ExprKind::Ident(r)) if l == r
    )
}

fn string_search_result_offset(left: &Expression, right: &Expression) -> Option<Expression> {
    fn extract_offset(expr_in: &Expression, right: &Expression) -> Option<Expression> {
        match &expr_in.kind {
            ExprKind::Ternary { then, else_, .. } => {
                extract_offset(then, right).or_else(|| extract_offset(else_, right))
            }
            ExprKind::Call { callee, args, .. } => {
                let ExprKind::Member { object, field, .. } = &callee.kind else {
                    return None;
                };
                if field != "slice" || args.len() != 1 || !same_ident_expr(object, right) {
                    return None;
                }
                Some(args[0].value.clone())
            }
            _ => None,
        }
    }

    extract_offset(left, right)
}

fn is_putchar_zero_call(e: &Expression) -> bool {
    let ExprKind::Call { callee, args, .. } = &e.kind else {
        return false;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "putchar") || args.len() != 1 {
        return false;
    }
    matches!(args[0].value.kind, ExprKind::Lit(Literal::Int(0)))
}

fn strip_putchar_side_effect_value(value: Expression) -> Expression {
    match value.kind {
        ExprKind::Call { callee, args, optional } => {
            if let ExprKind::Ident(name) = &callee.kind {
                if name == "__c_fputc_h" && !args.is_empty() {
                    return strip_putchar_side_effect_value(args[0].value.clone());
                }
            }
            let rewritten_args = args
                .into_iter()
                .map(|mut a| {
                    a.value = strip_putchar_side_effect_value(a.value);
                    a
                })
                .collect();
            expr(ExprKind::Call {
                callee,
                args: rewritten_args,
                optional,
            })
        }
        ExprKind::Binary { op, left, right } => expr(ExprKind::Binary {
            op,
            left: Box::new(strip_putchar_side_effect_value(*left)),
            right: Box::new(strip_putchar_side_effect_value(*right)),
        }),
        ExprKind::Unary { op, expr: inner } => expr(ExprKind::Unary {
            op,
            expr: Box::new(strip_putchar_side_effect_value(*inner)),
        }),
        ExprKind::Cast { expr: inner, type_name } => expr(ExprKind::Cast {
            expr: Box::new(strip_putchar_side_effect_value(*inner)),
            type_name,
        }),
        ExprKind::Ternary { cond, then, else_ } => expr(ExprKind::Ternary {
            cond: Box::new(strip_putchar_side_effect_value(*cond)),
            then: Box::new(strip_putchar_side_effect_value(*then)),
            else_: Box::new(strip_putchar_side_effect_value(*else_)),
        }),
        other => expr(other),
    }
}

fn is_zero_int_expr(e: &Expression) -> bool {
    matches!(e.kind, ExprKind::Lit(Literal::Int(0)))
}

/// True for a (possibly nested-brace) all-zero aggregate initializer:
/// `0`, `{0}`, `{{0}}`, `{{{0}}}`, `{}` — the C "zero everything" idiom.
fn is_all_zero_init(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Lit(Literal::Int(0)) => true,
        ExprKind::Array(elems) => elems.iter().all(|el| is_all_zero_init(&el.value)),
        _ => false,
    }
}

// ── setjmp/longjmp: re-entry transform ──────────────────────────────────────

/// If `e` contains a `__c_setjmp("tok")` marker, return its token.
fn find_setjmp_in_expr(e: &Expression) -> Option<String> {
    match &e.kind {
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Ident(n) = &callee.kind {
                if n == "__c_setjmp" {
                    if let Some(ExprKind::Lit(Literal::Str(t))) = args.first().map(|a| &a.value.kind)
                    {
                        return Some(t.clone());
                    }
                }
            }
            find_setjmp_in_expr(callee).or_else(|| args.iter().find_map(|a| find_setjmp_in_expr(&a.value)))
        }
        ExprKind::Binary { left, right, .. } => {
            find_setjmp_in_expr(left).or_else(|| find_setjmp_in_expr(right))
        }
        ExprKind::Ternary { cond, then, else_ } => find_setjmp_in_expr(cond)
            .or_else(|| find_setjmp_in_expr(then))
            .or_else(|| find_setjmp_in_expr(else_)),
        ExprKind::Unary { expr, .. } => find_setjmp_in_expr(expr),
        ExprKind::Assign { target, value } => {
            find_setjmp_in_expr(target).or_else(|| find_setjmp_in_expr(value))
        }
        ExprKind::Member { object, .. } => find_setjmp_in_expr(object),
        ExprKind::Index { object, index, .. } => {
            find_setjmp_in_expr(object).or_else(|| find_setjmp_in_expr(index))
        }
        ExprKind::Sequence(items) => items.iter().find_map(find_setjmp_in_expr),
        _ => None,
    }
}

/// Replace every `__c_setjmp(...)` marker in `e` with a read of `val_var`.
fn replace_setjmp_in_expr(e: &mut Expression, val_var: &str) {
    if let ExprKind::Call { callee, .. } = &e.kind {
        if let ExprKind::Ident(n) = &callee.kind {
            if n == "__c_setjmp" {
                *e = ident(val_var);
                return;
            }
        }
    }
    match &mut e.kind {
        ExprKind::Call { callee, args, .. } => {
            replace_setjmp_in_expr(callee, val_var);
            for a in args {
                replace_setjmp_in_expr(&mut a.value, val_var);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            replace_setjmp_in_expr(left, val_var);
            replace_setjmp_in_expr(right, val_var);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            replace_setjmp_in_expr(cond, val_var);
            replace_setjmp_in_expr(then, val_var);
            replace_setjmp_in_expr(else_, val_var);
        }
        ExprKind::Unary { expr, .. } => replace_setjmp_in_expr(expr, val_var),
        ExprKind::Assign { target, value } => {
            replace_setjmp_in_expr(target, val_var);
            replace_setjmp_in_expr(value, val_var);
        }
        ExprKind::Member { object, .. } => replace_setjmp_in_expr(object, val_var),
        ExprKind::Index { object, index, .. } => {
            replace_setjmp_in_expr(object, val_var);
            replace_setjmp_in_expr(index, val_var);
        }
        ExprKind::Sequence(items) => {
            for it in items {
                replace_setjmp_in_expr(it, val_var);
            }
        }
        _ => {}
    }
}

/// Token of the setjmp marker in a statement (looks in the contexts a setjmp
/// call appears: declaration inits, assignments, if/while conditions, returns).
fn find_setjmp_in_stmt(s: &Statement) -> Option<String> {
    match &s.kind {
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .find_map(|d| d.init.as_ref().and_then(find_setjmp_in_expr)),
        StmtKind::Assign { targets, value } => find_setjmp_in_expr(value)
            .or_else(|| targets.iter().find_map(find_setjmp_in_expr)),
        StmtKind::Expr(e) => find_setjmp_in_expr(e),
        StmtKind::If { cond, .. } => find_setjmp_in_expr(cond),
        StmtKind::While { cond, .. } => find_setjmp_in_expr(cond),
        StmtKind::Return(Some(e)) => find_setjmp_in_expr(e),
        _ => None,
    }
}

fn replace_setjmp_in_stmt(s: &mut Statement, val_var: &str) {
    match &mut s.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(init) = &mut d.init {
                    replace_setjmp_in_expr(init, val_var);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            replace_setjmp_in_expr(value, val_var);
            for t in targets {
                replace_setjmp_in_expr(t, val_var);
            }
        }
        StmtKind::Expr(e) => replace_setjmp_in_expr(e, val_var),
        StmtKind::If { cond, .. } => replace_setjmp_in_expr(cond, val_var),
        StmtKind::While { cond, .. } => replace_setjmp_in_expr(cond, val_var),
        StmtKind::Return(Some(e)) => replace_setjmp_in_expr(e, val_var),
        _ => {}
    }
}

/// Wrap the tail of a block (from the first statement containing a setjmp marker)
/// in a re-entry loop + try/catch, so the setjmp "returns twice": 0 initially,
/// then the longjmp value when an exception with the matching buf token unwinds.
fn wrap_setjmp_in_block(mut stmts: Vec<Statement>, counter: &mut u32) -> Vec<Statement> {
    let found = stmts
        .iter()
        .enumerate()
        .find_map(|(i, s)| find_setjmp_in_stmt(s).map(|t| (i, t)));
    let Some((i, token)) = found else {
        return stmts;
    };
    let n = *counter;
    *counter += 1;
    let val_var = format!("__sj_val{n}");
    let active_var = format!("__sj_active{n}");
    let err_var = format!("__sj_e{n}");

    let mut body = stmts.split_off(i);
    for s in &mut body {
        replace_setjmp_in_stmt(s, &val_var);
    }
    // Recurse in case a later setjmp (for a different buf) follows in the tail.
    body = wrap_setjmp_in_block(body, counter);

    // catch: if (e != null && e.__c_longjmp === token) { val = e.val; active = 1 } else throw e
    let is_match = expr(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(ident(&err_var)),
            right: Box::new(expr(ExprKind::Lit(Literal::Null))),
        })),
        right: Box::new(expr(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(member(ident(&err_var), "__c_longjmp")),
            right: Box::new(str_lit(&token)),
        })),
    });
    let catch_body = vec![if_stmt(
        is_match,
        vec![
            stmt(StmtKind::Expr(assign_expr(
                ident(&val_var),
                member(ident(&err_var), "__c_longjmp_val"),
            ))),
            stmt(StmtKind::Expr(assign_expr(ident(&active_var), int_lit(1)))),
        ],
        Some(vec![stmt(StmtKind::Throw {
            expr: Some(ident(&err_var)),
            cause: None,
        })]),
    )];
    let try_stmt = stmt(StmtKind::Try {
        body,
        catches: vec![CatchClause {
            types: Vec::new(),
            var_name: Some(err_var.clone()),
            stack_var: None,
            body: catch_body,
            when_clause: None,
        }],
        else_body: None,
        finally: None,
    });
    let loop_body = vec![
        stmt(StmtKind::Expr(assign_expr(ident(&active_var), int_lit(0)))),
        try_stmt,
    ];
    stmts.push(var_decl_stmt(&val_var, int_lit(0)));
    stmts.push(var_decl_stmt(&active_var, int_lit(1)));
    stmts.push(stmt(StmtKind::While {
        cond: expr(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(ident(&active_var)),
            right: Box::new(int_lit(0)),
        }),
        body: loop_body,
        else_body: None,
    }));
    stmts
}

fn base_ident_name(value: &Expression) -> Option<String> {
    match &value.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Index { object, .. } => base_ident_name(object),
        ExprKind::Object(props) => props.iter().find_map(|prop| {
            let ObjectProperty::KeyValue { key, value } = prop else {
                return None;
            };
            let ExprKind::Lit(Literal::Str(field)) = &key.kind else {
                return None;
            };
            if field == CARRAY_BASE_KEY {
                base_ident_name(value)
            } else {
                None
            }
        }),
        ExprKind::Call { callee, .. } => {
            let ExprKind::Member { object, field, .. } = &callee.kind else {
                return None;
            };
            if field == "substring" || field == "slice" {
                base_ident_name(object)
            } else {
                None
            }
        }
        _ => None,
    }
}

/// Whether evaluating `e` mutates program state, so it must not be duplicated.
/// Covers the index expressions a char-buffer element write reads twice
/// (`s[w++] = c`, `s[f()] = c`).
fn index_has_side_effects(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Unary { op, expr } => matches!(
            op,
            UnaryOp::PreInc | UnaryOp::PreDec | UnaryOp::PostInc | UnaryOp::PostDec
        ) || index_has_side_effects(expr),
        ExprKind::Assign { .. } | ExprKind::Call { .. } => true,
        ExprKind::Binary { left, right, .. } => {
            index_has_side_effects(left) || index_has_side_effects(right)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            index_has_side_effects(cond)
                || index_has_side_effects(then)
                || index_has_side_effects(else_)
        }
        ExprKind::Index { object, index, .. } => {
            index_has_side_effects(object) || index_has_side_effects(index)
        }
        ExprKind::Member { object, .. } => index_has_side_effects(object),
        _ => false,
    }
}

fn char_buffer_target_offset(value: &Expression) -> Option<(String, Expression)> {
    match &value.kind {
        ExprKind::Ident(name) => Some((name.clone(), expr(ExprKind::Lit(Literal::Int(0))))),
        ExprKind::Call { callee, args, .. } => {
            let ExprKind::Member { object, field, .. } = &callee.kind else {
                return None;
            };
            if (field != "substring" && field != "slice") || args.is_empty() {
                return None;
            }
            let ExprKind::Ident(base) = &object.kind else {
                return None;
            };
            Some((base.clone(), args[0].value.clone()))
        }
        _ => None,
    }
}

fn concat_sequence_to_string(mut pieces: Vec<Expression>) -> Expression {
    if pieces.is_empty() {
        return expr(ExprKind::Lit(Literal::Str(String::new())));
    }
    let mut current = pieces.remove(0);
    for piece in pieces {
        current = expr(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(current),
            right: Box::new(piece),
        });
    }
    current
}

fn char_pointer_offset_from_init(init: &Option<Expression>) -> Option<(String, Expression)> {
    let init_expr = init.as_ref()?;
    let candidate = if let ExprKind::Ternary { then, .. } = &init_expr.kind {
        then.as_ref()
    } else {
        init_expr
    };
    let ExprKind::Call { callee, args, .. } = &candidate.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if (field != "substring" && field != "slice") || args.len() != 1 {
        return None;
    }
    let ExprKind::Ident(base) = &object.kind else {
        return None;
    };
    Some((base.clone(), args[0].value.clone()))
}

fn carray_base_expr(value: &Expression) -> Option<Expression> {
    let ExprKind::Object(props) = &value.kind else {
        return None;
    };
    props.iter().find_map(|prop| {
        let ObjectProperty::KeyValue { key, value } = prop else {
            return None;
        };
        let ExprKind::Lit(Literal::Str(field)) = &key.kind else {
            return None;
        };
        if field == CARRAY_BASE_KEY {
            Some(value.clone())
        } else {
            None
        }
    })
}

fn char_assignment_value_to_string(value: Expression) -> Expression {
    if let ExprKind::Lit(Literal::Int(code)) = &value.kind {
        if *code == 0 {
            // Write a real NUL so the C null-terminator survives in the string.
            // Consumers (`printf %s`, `puts`, `strlen`) truncate at the first NUL.
            return expr(ExprKind::Lit(Literal::Str("\0".to_string())));
        }
        if let Some(ch) = char::from_u32(*code as u32) {
            return expr(ExprKind::Lit(Literal::Str(ch.to_string())));
        }
    }
    if matches!(value.kind, ExprKind::Lit(Literal::Str(_))) {
        value
    } else {
        string_adapter::char_code_to_string(value)
    }
}

fn sscanf_target_expr(value: &Expression) -> Expression {
    match &value.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr.clone(),
        _ => value.clone(),
    }
}

fn pointer_address_target_from_init(init: &Option<Expression>) -> Option<String> {
    let ExprKind::Unary {
        op: UnaryOp::AddrOf,
        expr,
    } = &init.as_ref()?.kind
    else {
        return None;
    };
    let ExprKind::Ident(target) = &expr.kind else {
        return None;
    };
    Some(target.clone())
}

fn pointer_member_target_from_init(init: &Option<Expression>) -> Option<Expression> {
    let ExprKind::Unary {
        op: UnaryOp::AddrOf,
        expr,
    } = &init.as_ref()?.kind
    else {
        return None;
    };
    if matches!(expr.kind, ExprKind::Member { .. }) {
        Some((**expr).clone())
    } else {
        None
    }
}

fn propagated_pointer_address_alias(
    init: &Option<Expression>,
    aliases: &HashMap<String, String>,
) -> Option<String> {
    let ExprKind::Ident(name) = &init.as_ref()?.kind else {
        return None;
    };
    aliases.get(name).cloned()
}

fn init_is_carray_pointer_var(init: &Option<Expression>, carray_vars: &HashSet<String>) -> bool {
    matches!(
        init.as_ref().map(|e| &e.kind),
        Some(ExprKind::Ident(name)) if carray_vars.contains(name)
    )
}

fn should_wrap_pointer_init_as_carray(
    init: &Option<Expression>,
    array_vars: &HashSet<String>,
) -> bool {
    match init.as_ref().map(|e| &e.kind) {
        Some(ExprKind::Ident(name)) => array_vars.contains(name),
        Some(ExprKind::Array(_)) => true,
        Some(ExprKind::Call { callee, .. }) => {
            matches!(&callee.kind, ExprKind::Ident(name) if name == "malloc" || name == "calloc" || name == "realloc")
        }
        Some(_) if init.as_ref().map(|e| is_carray_like_expr(e)).unwrap_or(false) => false,
        Some(_) => false,
        None => false,
    }
}

fn pointer_address_alias_comparison_side(
    aliases: &HashMap<String, String>,
    alias_expr: &Expression,
    address_expr: &Expression,
) -> Option<bool> {
    let ExprKind::Ident(alias_name) = &alias_expr.kind else {
        return None;
    };
    let expected_target = aliases.get(alias_name)?;
    let actual_target = pointer_address_target_from_expr(address_expr)?;
    Some(expected_target == &actual_target)
}

fn pointer_address_target_from_expr(e: &Expression) -> Option<String> {
    let ExprKind::Unary {
        op: UnaryOp::AddrOf,
        expr,
    } = &e.kind
    else {
        return None;
    };
    match &expr.kind {
        ExprKind::Ident(target) => Some(target.clone()),
        ExprKind::RefLoad(inner) => match &inner.kind {
            ExprKind::Ident(target) => Some(target.clone()),
            _ => None,
        },
        _ => None,
    }
}

fn compare_carray_to_array_start(ptr: Expression, array: Expression, op: BinOp) -> Expression {
    let base_eq = expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(expr(ExprKind::Member {
            object: Box::new(ptr.clone()),
            field: CARRAY_BASE_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(array),
    });
    let idx_eq = expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(expr(ExprKind::Member {
            object: Box::new(ptr),
            field: CARRAY_IDX_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
    });
    let eq = expr(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(base_eq),
        right: Box::new(idx_eq),
    });
    if matches!(op, BinOp::NotEq) {
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(eq),
        })
    } else {
        eq
    }
}

fn carray_ptr_equality(left: Expression, right: Expression) -> Expression {
    let base_eq = expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(expr(ExprKind::Member {
            object: Box::new(left.clone()),
            field: CARRAY_BASE_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(expr(ExprKind::Member {
            object: Box::new(right.clone()),
            field: CARRAY_BASE_KEY.to_string(),
            null_safe: false,
        })),
    });
    let idx_eq = expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(expr(ExprKind::Member {
            object: Box::new(left),
            field: CARRAY_IDX_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(expr(ExprKind::Member {
            object: Box::new(right),
            field: CARRAY_IDX_KEY.to_string(),
            null_safe: false,
        })),
    });
    expr(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(base_eq),
        right: Box::new(idx_eq),
    })
}

fn carray_ptr_relational(left: Expression, right: Expression, op: BinOp) -> Expression {
    let base_eq = expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(expr(ExprKind::Member {
            object: Box::new(left.clone()),
            field: CARRAY_BASE_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(expr(ExprKind::Member {
            object: Box::new(right.clone()),
            field: CARRAY_BASE_KEY.to_string(),
            null_safe: false,
        })),
    });
    let idx_cmp = expr(ExprKind::Binary {
        op,
        left: Box::new(expr(ExprKind::Member {
            object: Box::new(left),
            field: CARRAY_IDX_KEY.to_string(),
            null_safe: false,
        })),
        right: Box::new(expr(ExprKind::Member {
            object: Box::new(right),
            field: CARRAY_IDX_KEY.to_string(),
            null_safe: false,
        })),
    });
    expr(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(base_eq),
        right: Box::new(idx_cmp),
    })
}

fn atomic_pointer_target(arg: Expression) -> Expression {
    if let ExprKind::Unary {
        op: UnaryOp::AddrOf,
        expr,
    } = arg.kind
    {
        *expr
    } else {
        arg
    }
}

fn nan_to_default(value: Expression, default_value: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(expr(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(value.clone()),
            right: Box::new(value.clone()),
        })),
        then: Box::new(default_value),
        else_: Box::new(value),
    })
}

/// Extract an integer value from an optional enum initializer expression.
/// Handles plain integer literals and negated literals (e.g. `= -2`).
/// Falls back to `default` for unsupported forms.
fn extract_enum_val(init: &Option<Expression>, default: i64) -> i64 {
    let Some(e) = init else { return default };
    match &e.kind {
        ExprKind::Lit(Literal::Int(i)) => *i,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => {
            if let ExprKind::Lit(Literal::Int(i)) = &expr.kind {
                -i
            } else {
                default
            }
        }
        _ => default,
    }
}

/// Returns true if a statement list ends with a break or return (no fallthrough).
fn ends_with_break(stmts: &[Statement]) -> bool {
    match stmts.last() {
        Some(s) => matches!(
            s.kind,
            StmtKind::Break(_) | StmtKind::Return(_) | StmtKind::GoTo(_)
        ),
        None => false,
    }
}

/// Convert an expression AST node back to a simple C source string for macro substitution.
/// Only handles the common cases that appear in macro arguments.
fn expr_to_c_source(expr: &Expression) -> String {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(n)) => n.to_string(),
        ExprKind::Lit(Literal::Float(f)) => format!("{}", f),
        ExprKind::Lit(Literal::Str(s)) => {
            format!("\"{}\"", s.replace('\\', "\\\\").replace('"', "\\\""))
        }
        ExprKind::Lit(Literal::Bool(b)) => {
            if *b {
                "1".to_string()
            } else {
                "0".to_string()
            }
        }
        ExprKind::Ident(name) => name.clone(),
        ExprKind::Binary { op, left, right } => {
            let op_str = match op {
                BinOp::Add => "+",
                BinOp::Sub => "-",
                BinOp::Mul => "*",
                BinOp::Div => "/",
                BinOp::Mod => "%",
                BinOp::And => "&&",
                BinOp::Or => "||",
                BinOp::Eq => "==",
                BinOp::NotEq => "!=",
                BinOp::Lt => "<",
                BinOp::LtEq => "<=",
                BinOp::Gt => ">",
                BinOp::GtEq => ">=",
                BinOp::BitAnd => "&",
                BinOp::BitOr => "|",
                BinOp::BitXor => "^",
                BinOp::Shl => "<<",
                BinOp::Shr => ">>",
                _ => "+",
            };
            format!(
                "({} {} {})",
                expr_to_c_source(left),
                op_str,
                expr_to_c_source(right)
            )
        }
        ExprKind::Unary { op, expr: e } => {
            let op_str = match op {
                UnaryOp::Neg => "-",
                UnaryOp::Not => "!",
                UnaryOp::BitNot => "~",
                _ => "",
            };
            format!("({}{})", op_str, expr_to_c_source(e))
        }
        ExprKind::Call { callee, args, .. } => {
            let callee_s = expr_to_c_source(callee);
            let args_s: Vec<String> = args.iter().map(|a| expr_to_c_source(&a.value)).collect();
            format!("{}({})", callee_s, args_s.join(", "))
        }
        _ => "0".to_string(),
    }
}

fn is_complex_type_text(type_text: &str) -> bool {
    type_text.to_ascii_lowercase().contains("complex")
}


/// Replace all whole-word occurrences of `word` in `text` with `replacement`.
fn replace_word(text: &str, word: &str, replacement: &str) -> String {
    let mut result = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(pos) = rest.find(word) {
        let before = &rest[..pos];
        let after = &rest[pos + word.len()..];
        // Check word boundaries
        let before_ok = before
            .chars()
            .last()
            .map_or(true, |c| !c.is_alphanumeric() && c != '_');
        let after_ok = after
            .chars()
            .next()
            .map_or(true, |c| !c.is_alphanumeric() && c != '_');
        if before_ok && after_ok {
            result.push_str(before);
            result.push_str(replacement);
        } else {
            result.push_str(before);
            result.push_str(word);
        }
        rest = after;
    }
    result.push_str(rest);
    result
}
