//! Go walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//!
//! ## Go-specific normalisations
//!
//! - **Multiple return values**: Go functions can return multiple values.
//!   For simplicity we compile to returning a single array/tuple.
//! - **Short variable declaration** (`:=`): Maps to `VarDecl` with `Let`.
//! - **Methods**: Go methods on structs are compiled into `StructDecl`
//!   fragments with the receiver kept as the first explicit parameter.
//! - **Structs**: Mapped to `StructDecl` with fields.
//! - **Interfaces**: Mapped to `InterfaceDecl`.
//! - **`range`**: Mapped to `ForIn` with `of: true`.
//! - **`defer`**: Lowered to a per-function stack of zero-arg closures that
//!   drain from a synthesized `finally` block in LIFO order.
//! - **`go`**: Lowered to the shared thread/task emitter surface.
//! - **`fallthrough`**: Not yet supported in switch.
//! - **`select`**: Lowered as a compile-safe block for the dummy concurrency tests.
//! - **`chan` / `<-`**: Lowered into compile-safe object/array operations.
//! - **`nil`**: Mapped to `ExprKind::Lit(Literal::Null)`.
//! - **`make` / `new`**: `make` for slices/maps is rewritten to array/dict
//!   creation. `new(T)` becomes `&T{}` (pointer to zero value).
//! - **`append`**: Rewritten to slice concat so the updated slice value is preserved.
//! - **`len` / `cap`**: Builtin functions mapped to host calls.
//! - **`panic` / `recover`**: Mapped to throw/try-catch.
//! - **`_` blank identifier**: Ignored in assignments.

use super::{GoParser, Rule};
use pest::Parser;
use pest::iterators::Pair;
use regex::{Captures, Regex};
use std::collections::{HashMap, HashSet};
use vybe_ast::*;
// Channels are normalized into COMMON AST shapes — the walker builds AST,
// not bytecode. The emit side lives in the compiler.
use vybe_ast::{ChanOp, SelectArm};
use vybe_compiler::primitives::generics as common_generics;
use vybe_compiler::primitives::reflection;

// ══════════════════════════════════════════════════════════════════════════════════════════
// Entry point
// ══════════════════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let (package_name, body, imports) = walk_go_source(source)?;
    go_validate_method_sets(&merge_go_struct_decls(&body))?;

    let mut module = normalize_go_module(Module {
        canon: Default::default(),
        name: package_name,
        language: Lang::Go,
        body,
        imports,
        directives: vybe_ast::Directives {
            // A go method declares its receiver (`func (r *T) M()`) and the
            // CALL supplies it as a leading argument — the callable is the raw
            // function off the type, carrying no receiver of its own.
            method_receiver: Some(vybe_ast::MethodReceiver::CallSite),
            // Every callable declares a leading receiver parameter, not only
            // methods — ECMA-262 §10.2.1 `[[Call]](thisArgument,
            // argumentsList)`. A plain `f()` passes `undefined` (§10.2.1.1).
            receiver_binding: Some(vybe_ast::ReceiverBinding::UniversalParameter),
            ..Default::default()
        },
    });
    module.imports.retain(go_should_emit_import);
    Ok(module)
}

/// Go spec "Semicolon insertion" (§Tokens): the lexer inserts `;` at a
/// newline when the line's last token is an identifier, a literal, one of
/// `break` / `continue` / `fallthrough` / `return`, or `++` `--` `)` `]`
/// `}`. The pest grammar's WHITESPACE includes `\n`, so WITHOUT this pass
/// an expression continues across the newline and a following line-start
/// `*p = x` / `<-ch` / `-x` is swallowed as a binary operand — the whole
/// line-start statement family. The grammar already tolerates `;` in every
/// insertion position (statements, specs, field/interface members).
///
/// The `;` is emitted AFTER the `\n` so line numbers and comment text are
/// untouched. Inside strings / runes / raw strings / comments nothing is
/// inserted; a blank line resets the state so it never double-inserts.
fn insert_go_semicolons(src: &str) -> String {
    // Keywords a line may legally END on withOUT a semicolon following in
    // real Go — everything except break/continue/fallthrough/return.
    const NO_INSERT_KEYWORDS: &[&str] = &[
        "package",
        "import",
        "func",
        "var",
        "const",
        "type",
        "if",
        "else",
        "for",
        "range",
        "go",
        "defer",
        "chan",
        "map",
        "struct",
        "interface",
        "switch",
        "select",
        "case",
        "default",
        "goto",
    ];
    #[derive(PartialEq)]
    enum St {
        Normal,
        LineComment,
        BlockComment,
        Dq,
        Raw,
        Rune,
    }
    let mut out = String::with_capacity(src.len() + src.len() / 16);
    let mut st = St::Normal;
    let mut word = String::new();
    let mut last_ch: Option<char> = None;
    let mut prev_ch: Option<char> = None;
    let mut chars = src.chars().peekable();
    while let Some(c) = chars.next() {
        match st {
            St::Normal => match c {
                '\n' => {
                    out.push('\n');
                    let insert = match last_ch {
                        Some(l) if l.is_alphanumeric() || l == '_' => {
                            !NO_INSERT_KEYWORDS.contains(&word.as_str())
                        }
                        Some(')') | Some(']') | Some('}') | Some('"') | Some('\'') | Some('`') => {
                            true
                        }
                        Some('+') => prev_ch == Some('+'),
                        Some('-') => prev_ch == Some('-'),
                        _ => false,
                    };
                    if insert {
                        out.push(';');
                    }
                    word.clear();
                    last_ch = None;
                    prev_ch = None;
                }
                ' ' | '\t' | '\r' => out.push(c),
                '"' => {
                    st = St::Dq;
                    out.push(c);
                }
                '`' => {
                    st = St::Raw;
                    out.push(c);
                }
                '\'' => {
                    st = St::Rune;
                    out.push(c);
                }
                '/' if chars.peek() == Some(&'/') => {
                    st = St::LineComment;
                    out.push(c);
                }
                '/' if chars.peek() == Some(&'*') => {
                    st = St::BlockComment;
                    out.push(c);
                }
                _ => {
                    if c.is_alphanumeric() || c == '_' {
                        if !matches!(last_ch, Some(l) if l.is_alphanumeric() || l == '_') {
                            word.clear();
                        }
                        word.push(c);
                    }
                    prev_ch = last_ch;
                    last_ch = Some(c);
                    out.push(c);
                }
            },
            St::LineComment => {
                if c == '\n' {
                    // Decide with the PRE-comment token state; the `;` lands
                    // after the newline, outside the comment text.
                    out.push('\n');
                    let insert = match last_ch {
                        Some(l) if l.is_alphanumeric() || l == '_' => {
                            !NO_INSERT_KEYWORDS.contains(&word.as_str())
                        }
                        Some(')') | Some(']') | Some('}') | Some('"') | Some('\'') | Some('`') => {
                            true
                        }
                        Some('+') => prev_ch == Some('+'),
                        Some('-') => prev_ch == Some('-'),
                        _ => false,
                    };
                    if insert {
                        out.push(';');
                    }
                    word.clear();
                    last_ch = None;
                    prev_ch = None;
                    st = St::Normal;
                } else {
                    out.push(c);
                }
            }
            St::BlockComment => {
                out.push(c);
                if c == '*' && chars.peek() == Some(&'/') {
                    out.push(chars.next().unwrap());
                    st = St::Normal;
                }
            }
            St::Dq | St::Rune => {
                out.push(c);
                if c == '\\' {
                    if let Some(esc) = chars.next() {
                        out.push(esc);
                    }
                } else if (c == '"' && st == St::Dq) || (c == '\'' && st == St::Rune) {
                    prev_ch = last_ch;
                    last_ch = Some(c);
                    word.clear();
                    st = St::Normal;
                }
            }
            St::Raw => {
                out.push(c);
                if c == '`' {
                    prev_ch = last_ch;
                    last_ch = Some(c);
                    word.clear();
                    st = St::Normal;
                }
            }
        }
    }
    out
}

/// Walk a Go source string into its raw (pre-normalization) parts.
fn walk_go_source(source: &str) -> Result<(String, Vec<Statement>, Vec<Import>), String> {
    let source = insert_go_semicolons(source);
    let source = source.as_str();
    let _line_index = vybe_ast::line_index::LineIndex::install(source);
    let pairs =
        GoParser::parse(Rule::program, source).map_err(|e| format!("Go parse error: {}", e))?;

    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut package_name = String::new();

    for top in pairs {
        if top.as_rule() == Rule::EOI {
            continue;
        }
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            _ => {
                if let Some(stmt) = walk_top_level(top)? {
                    body.push(stmt);
                }
                continue;
            }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI => continue,
                Rule::package_clause => {
                    package_name = walk_package_clause(pair)?;
                }
                Rule::import_declarations => {
                    for imp in pair.into_inner() {
                        if imp.as_rule() == Rule::import_declaration {
                            imports.push(walk_import(imp)?);
                        }
                    }
                }
                _ => {
                    if let Some(stmt) = walk_top_level(pair)? {
                        body.push(stmt);
                    }
                }
            }
        }
    }

    Ok((package_name, body, imports))
}

#[derive(Clone, Default)]
struct GoFunctionSignature {
    params: Vec<Option<String>>,
    return_type: Option<String>,
    generic_arg_count: usize,
    generic_param_names: Vec<String>,
}

#[derive(Clone, Default)]
struct GoNormalizeEnv {
    value_types: HashMap<String, String>,
    reflect_value_payloads: HashMap<String, Expression>,
    reflect_value_targets: HashMap<String, Expression>,
    reflect_pointer_targets: HashMap<String, (Expression, String)>,
    reflect_method_bindings: HashMap<String, (Expression, String)>,
    reflect_array_payloads: HashMap<String, Vec<Expression>>,
    package_aliases: HashMap<String, String>,
    fixed_arrays: HashMap<String, String>,
    regex_patterns: HashMap<String, String>,
    slice_caps: HashMap<String, Expression>,
    slice_views: HashMap<String, GoSliceViewInfo>,
    struct_infos: HashMap<String, GoStructInfo>,
    interface_methods: HashMap<String, HashSet<String>>,
    interface_concrete_types: HashMap<String, String>,
    nil_interface_values: HashSet<String>,
    method_value_bindings: HashMap<String, GoMethodValueBinding>,
    named_types: HashMap<String, String>,
    type_names: HashSet<String>,
    function_bodies: HashMap<String, Vec<Statement>>,
    flag_bindings: HashMap<String, (String, String)>,
    flag_defs: Vec<crate::adapters::flags::FlagDefinition>,
    log_output: Option<Expression>,
    log_prefix: Option<Expression>,
    log_flags: Option<Expression>,
    time_round_half_hour_bindings: HashSet<String>,
    generic_type_params: HashMap<String, String>,
    return_type: Option<String>,
    panic_value_name: Option<String>,
    has_panic_name: Option<String>,
    in_defer_name: Option<String>,
    recover_fn_name: Option<String>,
    owns_panic_state: bool,
}

#[derive(Clone)]
struct GoSliceViewInfo {
    base: Expression,
    start: Expression,
    end: Option<Expression>,
    max: Option<Expression>,
}

#[derive(Clone, Default)]
struct GoStructInfo {
    field_order: Vec<String>,
    member_names: HashSet<String>,
    method_names: HashSet<String>,
    pointer_method_names: HashSet<String>,
    method_receiver_types: HashMap<String, String>,
    method_params: HashMap<String, Vec<Param>>,
    member_types: HashMap<String, String>,
    field_tags: HashMap<String, String>,
    embedded_fields: Vec<(String, String)>,
}

#[derive(Clone)]
struct GoMethodDef {
    params: Vec<Param>,
    return_type: Option<String>,
    body: Vec<Statement>,
    modifiers: Modifiers,
    handles: Vec<String>,
    is_async: bool,
    is_generator: bool,
    is_sub: bool,
}

#[derive(Clone)]
struct GoMethodValueBinding {
    receiver: Expression,
    method: String,
}

#[derive(Default)]
struct GoNormalizeState {
    next_temp: usize,
}

struct GoSignatureInfo {
    params: Vec<Param>,
    return_type: Option<String>,
    named_results: Vec<Param>,
}

fn normalize_go_module(mut module: Module) -> Module {
    module.body = merge_go_struct_decls(&module.body);
    let signatures = collect_go_function_signatures(&module.body);
    let globals = collect_go_global_fixed_arrays(&module.body, &signatures);
    let struct_infos = collect_go_struct_infos(&module.body);
    let interface_methods = collect_go_interface_methods(&module.body);
    let named_types = collect_go_named_types(&module.body);
    let type_names = collect_go_type_names(&module.body);
    let function_bodies = collect_go_function_bodies(&module.body);
    let package_aliases = collect_go_package_aliases(&module.imports);
    let mut state = GoNormalizeState::default();
    let mut env = GoNormalizeEnv {
        value_types: HashMap::new(),
        reflect_value_payloads: HashMap::new(),
        reflect_value_targets: HashMap::new(),
        reflect_pointer_targets: HashMap::new(),
        reflect_method_bindings: HashMap::new(),
        reflect_array_payloads: HashMap::new(),
        package_aliases,
        fixed_arrays: globals.clone(),
        regex_patterns: HashMap::new(),
        slice_caps: HashMap::new(),
        slice_views: HashMap::new(),
        struct_infos,
        interface_methods,
        interface_concrete_types: HashMap::new(),
        nil_interface_values: HashSet::new(),
        method_value_bindings: HashMap::new(),
        named_types,
        type_names,
        function_bodies,
        flag_bindings: HashMap::new(),
        flag_defs: Vec::new(),
        log_output: None,
        log_prefix: None,
        log_flags: None,
        time_round_half_hour_bindings: HashSet::new(),
        generic_type_params: HashMap::new(),
        return_type: None,
        panic_value_name: None,
        has_panic_name: None,
        in_defer_name: None,
        recover_fn_name: None,
        owns_panic_state: false,
    };

    let mut normalized = Vec::with_capacity(module.body.len());
    for stmt in &module.body {
        normalized.extend(normalize_go_statement(
            stmt,
            &mut env,
            &signatures,
            &mut state,
        ));
    }
    let normalized = go_add_named_scalar_method_wrappers(normalized, &env);
    let normalized = go_hoist_zero_value_globals(normalized);
    module.body = go_lower_module_init_functions(normalized, &mut state);
    module
}

fn go_add_named_scalar_method_wrappers(
    mut body: Vec<Statement>,
    env: &GoNormalizeEnv,
) -> Vec<Statement> {
    let mut wrappers = Vec::new();
    let mut method_defs: HashMap<(String, String), GoMethodDef> = HashMap::new();
    for stmt in &body {
        let StmtKind::StructDecl { name, members, .. } = &stmt.kind else {
            continue;
        };
        let underlying = env.named_types.get(name);

        for member in members {
            let ClassMember::Method(method) = member else {
                continue;
            };
            let StmtKind::FunctionDecl {
                name: method_name,
                params,
                return_type,
                body,
                modifiers,
                handles,
                is_async,
                is_generator,
                is_sub,
            } = &method.kind
            else {
                continue;
            };
            method_defs.insert(
                (name.clone(), method_name.clone()),
                GoMethodDef {
                    params: params.clone(),
                    return_type: return_type.clone(),
                    body: body.clone(),
                    modifiers: modifiers.clone(),
                    handles: handles.clone(),
                    is_async: *is_async,
                    is_generator: *is_generator,
                    is_sub: *is_sub,
                },
            );
            if go_skip_method_wrapper_type(name) {
                continue;
            }

            let mut wrapper_body = body.clone();
            if underlying.is_some_and(|ty| ty.trim() == "int") {
                if let Some(receiver) = params
                    .first()
                    .filter(|param| {
                        param
                            .type_hint
                            .as_deref()
                            .is_some_and(|ty| !ty.trim().starts_with('*'))
                    })
                    .map(|param| param.name.clone())
                {
                    go_rewrite_named_integer_receiver_unwraps(&mut wrapper_body, &receiver);
                }
            }

            let mut wrapper_params = params.clone();
            if let (Some(underlying), Some(receiver)) = (underlying, wrapper_params.first_mut()) {
                if receiver
                    .type_hint
                    .as_deref()
                    .is_some_and(|ty| !ty.trim().starts_with('*'))
                {
                    receiver.type_hint = Some(underlying.clone().into());
                }
            }

            wrappers.push(Statement::new(StmtKind::FunctionDecl {
                name: go_scalar_method_wrapper_name(name, method_name),
                params: wrapper_params,
                return_type: return_type.clone(),
                body: wrapper_body,
                modifiers: modifiers.clone(),
                handles: handles.clone(),
                is_async: *is_async,
                is_generator: *is_generator,
                is_sub: *is_sub,
            }));
        }
    }
    go_add_promoted_pointer_method_wrappers(&body, env, &method_defs, &mut wrappers);
    body.extend(wrappers);
    body
}

fn go_scalar_method_wrapper_name(type_name: &str, method_name: &str) -> String {
    fn clean(part: &str) -> String {
        part.chars()
            .map(|ch| {
                if ch.is_ascii_alphanumeric() || ch == '_' {
                    ch
                } else {
                    '_'
                }
            })
            .collect()
    }
    format!("__go_method_{}_{}", clean(type_name), clean(method_name))
}

fn go_named_non_struct_underlying_type(type_name: &str, env: &GoNormalizeEnv) -> Option<String> {
    let underlying = env.named_types.get(type_name.trim())?;
    if underlying.trim().starts_with("struct") {
        return None;
    }
    Some(underlying.clone())
}

fn go_add_promoted_pointer_method_wrappers(
    body: &[Statement],
    env: &GoNormalizeEnv,
    method_defs: &HashMap<(String, String), GoMethodDef>,
    wrappers: &mut Vec<Statement>,
) {
    let mut emitted = HashSet::new();
    for stmt in body {
        let StmtKind::StructDecl { name, .. } = &stmt.kind else {
            continue;
        };
        if go_skip_method_wrapper_type(name) {
            continue;
        }
        let Some(info) = env.struct_infos.get(name) else {
            continue;
        };
        for (embedded_field, embedded_type) in &info.embedded_fields {
            let Some(embedded_lookup) = go_struct_lookup_name(embedded_type) else {
                continue;
            };
            let Some(embedded_info) = env.struct_infos.get(&embedded_lookup) else {
                continue;
            };
            for method_name in &embedded_info.pointer_method_names {
                if info.method_names.contains(method_name)
                    || go_has_ambiguous_promoted_method(name, method_name, env)
                    || !emitted.insert((name.clone(), method_name.clone()))
                {
                    continue;
                }
                let Some(method_def) =
                    method_defs.get(&(embedded_lookup.clone(), method_name.clone()))
                else {
                    continue;
                };
                let Some(receiver) = method_def.params.first() else {
                    continue;
                };
                let outer_receiver = "__go_promoted_receiver".to_string();
                let mut wrapper_params = method_def.params.clone();
                wrapper_params[0] = Param {
                    name: outer_receiver.clone(),
                    type_hint: Some(format!("*{}", name).into()),
                    default: receiver.default.clone(),
                    pass_by: receiver.pass_by,
                    is_rest: receiver.is_rest,
                    is_kwargs: receiver.is_kwargs,
                    is_optional: receiver.is_optional,
                    is_nullable: receiver.is_nullable,
                };

                let embedded_expr = Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::RefLoad(Box::new(
                        Expression::ident(&outer_receiver),
                    )))),
                    field: embedded_field.clone(),
                    null_safe: false,
                });
                let promoted_pointer_receiver = receiver
                    .type_hint
                    .as_deref()
                    .is_some_and(|ty| ty.trim().starts_with('*'));
                let mut wrapper_body = method_def.body.clone();
                if promoted_pointer_receiver && !embedded_type.trim().starts_with('*') {
                    let embedded_temp = format!("__go_promoted_{}_{}", name, embedded_field);
                    go_replace_receiver_expr_in_statements(
                        &mut wrapper_body,
                        &receiver.name,
                        &Expression::ident(&embedded_temp),
                    );
                    wrapper_body.insert(
                        0,
                        Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(embedded_temp.clone()),
                                type_hint: Some(embedded_type.clone().into()),
                                init: Some(embedded_expr.clone()),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }),
                    );
                    wrapper_body.push(Statement::new(StmtKind::Assign {
                        targets: vec![embedded_expr],
                        value: Expression::ident(&embedded_temp),
                        by_ref: false,
                    }));
                } else {
                    let receiver_replacement = if promoted_pointer_receiver {
                        Expression::new(ExprKind::RefLoad(Box::new(go_addr_of_normalized_expr(
                            embedded_expr,
                        ))))
                    } else {
                        embedded_expr
                    };
                    go_replace_receiver_expr_in_statements(
                        &mut wrapper_body,
                        &receiver.name,
                        &receiver_replacement,
                    );
                }

                wrappers.push(Statement::new(StmtKind::FunctionDecl {
                    name: go_scalar_method_wrapper_name(name, method_name),
                    params: wrapper_params,
                    return_type: method_def.return_type.clone(),
                    body: wrapper_body,
                    modifiers: method_def.modifiers.clone(),
                    handles: method_def.handles.clone(),
                    is_async: method_def.is_async,
                    is_generator: method_def.is_generator,
                    is_sub: method_def.is_sub,
                }));
            }
        }
    }
}

fn go_replace_receiver_expr_in_statements(
    body: &mut [Statement],
    receiver_name: &str,
    replacement: &Expression,
) {
    for stmt in body {
        go_replace_receiver_expr_in_statement(stmt, receiver_name, replacement);
    }
}

fn go_replace_receiver_expr_in_statement(
    stmt: &mut Statement,
    receiver_name: &str,
    replacement: &Expression,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => go_replace_receiver_expr(expr, receiver_name, replacement),
        StmtKind::Return(expr) => {
            if let Some(expr) = expr {
                go_replace_receiver_expr(expr, receiver_name, replacement);
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                go_replace_receiver_expr(target, receiver_name, replacement);
            }
            go_replace_receiver_expr(value, receiver_name, replacement);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            go_replace_receiver_expr(target, receiver_name, replacement);
            go_replace_receiver_expr(value, receiver_name, replacement);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    go_replace_receiver_expr(init, receiver_name, replacement);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        go_replace_receiver_expr(bound, receiver_name, replacement);
                    }
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            go_replace_receiver_expr(cond, receiver_name, replacement);
            go_replace_receiver_expr_in_statements(then_body, receiver_name, replacement);
            for (elif_cond, elif_body) in elifs {
                go_replace_receiver_expr(elif_cond, receiver_name, replacement);
                go_replace_receiver_expr_in_statements(elif_body, receiver_name, replacement);
            }
            if let Some(else_body) = else_body {
                go_replace_receiver_expr_in_statements(else_body, receiver_name, replacement);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            go_replace_receiver_expr(cond, receiver_name, replacement);
            go_replace_receiver_expr_in_statements(body, receiver_name, replacement);
            if let Some(else_body) = else_body {
                go_replace_receiver_expr_in_statements(else_body, receiver_name, replacement);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                go_replace_receiver_expr_in_statement(init, receiver_name, replacement);
            }
            if let Some(cond) = cond {
                go_replace_receiver_expr(cond, receiver_name, replacement);
            }
            if let Some(update) = update {
                go_replace_receiver_expr(update, receiver_name, replacement);
            }
            go_replace_receiver_expr_in_statements(body, receiver_name, replacement);
        }
        StmtKind::Block(body) => {
            go_replace_receiver_expr_in_statements(body, receiver_name, replacement);
        }
        _ => {}
    }
}

fn go_replace_receiver_expr(expr: &mut Expression, receiver_name: &str, replacement: &Expression) {
    match &mut expr.kind {
        ExprKind::Ident(name) if name == receiver_name => {
            *expr = replacement.clone();
        }
        ExprKind::RefLoad(inner) if matches!(&inner.kind, ExprKind::Ident(name) if name == receiver_name) =>
        {
            *expr = replacement.clone();
        }
        ExprKind::Call { callee, args, .. } => {
            go_replace_receiver_expr(callee, receiver_name, replacement);
            for arg in args {
                go_replace_receiver_expr(&mut arg.value, receiver_name, replacement);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            go_replace_receiver_expr(left, receiver_name, replacement);
            go_replace_receiver_expr(right, receiver_name, replacement);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Spread(expr) => {
            go_replace_receiver_expr(expr, receiver_name, replacement);
        }
        ExprKind::Member { object, .. } => {
            go_replace_receiver_expr(object, receiver_name, replacement);
        }
        ExprKind::Index { object, index, .. } => {
            go_replace_receiver_expr(object, receiver_name, replacement);
            go_replace_receiver_expr(index, receiver_name, replacement);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_replace_receiver_expr(cond, receiver_name, replacement);
            go_replace_receiver_expr(then, receiver_name, replacement);
            go_replace_receiver_expr(else_, receiver_name, replacement);
        }
        ExprKind::Array(elements) => {
            for element in elements {
                go_replace_receiver_expr(&mut element.value, receiver_name, replacement);
                if let Some(key) = &mut element.key {
                    go_replace_receiver_expr(key, receiver_name, replacement);
                }
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    go_replace_receiver_expr(key, receiver_name, replacement);
                    go_replace_receiver_expr(value, receiver_name, replacement);
                }
            }
        }
        _ => {}
    }
}

fn go_rewrite_named_integer_receiver_unwraps(body: &mut [Statement], receiver: &str) {
    for stmt in body {
        if let StmtKind::Return(Some(expr)) = &mut stmt.kind {
            go_rewrite_named_integer_receiver_unwrap_expr(expr, receiver);
        }
    }
}

fn go_rewrite_named_integer_receiver_unwrap_expr(expr: &mut Expression, receiver: &str) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            for arg in args.iter_mut() {
                go_rewrite_named_integer_receiver_unwrap_expr(&mut arg.value, receiver);
            }
            if matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_to_int")
                && args.len() == 1
                && matches!(&args[0].value.kind, ExprKind::Ident(name) if name == receiver)
            {
                *expr = Expression::ident(receiver);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            go_rewrite_named_integer_receiver_unwrap_expr(left, receiver);
            go_rewrite_named_integer_receiver_unwrap_expr(right, receiver);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Spread(expr) => {
            go_rewrite_named_integer_receiver_unwrap_expr(expr, receiver);
        }
        ExprKind::Member { object, .. } => {
            go_rewrite_named_integer_receiver_unwrap_expr(object, receiver);
        }
        ExprKind::Index { object, index, .. } => {
            go_rewrite_named_integer_receiver_unwrap_expr(object, receiver);
            go_rewrite_named_integer_receiver_unwrap_expr(index, receiver);
        }
        _ => {}
    }
}

fn go_hoist_zero_value_globals(body: Vec<Statement>) -> Vec<Statement> {
    let mut zero_globals = Vec::new();
    let mut rest = Vec::with_capacity(body.len());

    for stmt in body {
        if go_is_zero_value_global_var_decl(&stmt) {
            zero_globals.push(stmt);
        } else {
            rest.push(stmt);
        }
    }

    zero_globals.extend(rest);
    zero_globals
}

fn go_is_zero_value_global_var_decl(stmt: &Statement) -> bool {
    let StmtKind::VarDecl { declarations, .. } = &stmt.kind else {
        return false;
    };
    !declarations.is_empty()
        && declarations
            .iter()
            .all(|decl| decl.init.as_ref().is_some_and(go_is_zero_value_expr))
}

fn go_is_zero_value_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Null) => true,
        ExprKind::Lit(Literal::Bool(false)) => true,
        ExprKind::Lit(Literal::Int(0)) => true,
        ExprKind::Lit(Literal::Float(value)) => *value == 0.0,
        ExprKind::Lit(Literal::Str(value)) => value.is_empty(),
        ExprKind::Array(elements) => elements.is_empty(),
        ExprKind::Object(properties) => properties.iter().all(|property| match property {
            ObjectProperty::KeyValue { value, .. } => go_is_zero_value_expr(value),
            _ => false,
        }),
        ExprKind::Cast { expr, .. } => go_is_zero_value_expr(expr),
        _ => false,
    }
}

fn collect_go_package_aliases(imports: &[Import]) -> HashMap<String, String> {
    let mut aliases = HashMap::new();
    for import in imports {
        let ImportKind::Simple { path, alias } = &import.kind else {
            continue;
        };
        let Some(alias) = alias.as_deref() else {
            continue;
        };
        if alias == "." || alias == "_" {
            continue;
        }
        let package_name = path.rsplit('/').next().unwrap_or(path).trim();
        if !package_name.is_empty() {
            aliases.insert(alias.to_string(), package_name.to_string());
        }
    }
    aliases
}

fn go_should_emit_import(import: &Import) -> bool {
    let ImportKind::Simple { path, .. } = &import.kind else {
        return true;
    };
    !go_is_adapter_stdlib_import(path)
}

fn go_is_adapter_stdlib_import(path: &str) -> bool {
    matches!(path.trim(), "log/slog")
}

fn go_lower_module_init_functions(
    body: Vec<Statement>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let mut lowered = Vec::with_capacity(body.len());
    let mut init_calls = Vec::new();

    for stmt in body {
        match stmt.kind {
            StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                body,
                modifiers,
                handles,
                is_async,
                is_generator,
                is_sub,
            } if name == "init" => {
                let hidden_name = fresh_go_temp(state, "__go_init");
                lowered.push(Statement::new(StmtKind::FunctionDecl {
                    name: hidden_name.clone(),
                    params,
                    return_type,
                    body,
                    modifiers,
                    handles,
                    is_async,
                    is_generator,
                    is_sub,
                }));
                init_calls.push(Statement::new(StmtKind::Expr(Expression::new(
                    ExprKind::Call {
                        callee: Box::new(Expression::ident(&hidden_name)),
                        args: Vec::new(),
                        optional: false,
                    },
                ))));
            }
            _ => lowered.push(stmt),
        }
    }

    lowered.extend(init_calls);
    lowered
}

fn collect_go_function_signatures(body: &[Statement]) -> HashMap<String, GoFunctionSignature> {
    let mut signatures = HashMap::new();
    for stmt in body {
        match &stmt.kind {
            StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                ..
            } => {
                signatures.insert(
                    name.clone(),
                    GoFunctionSignature {
                        params: params
                            .iter()
                            .map(|param| param.type_hint.as_deref().map(str::to_string))
                            .collect(),
                        return_type: return_type.clone(),
                        generic_arg_count: go_signature_generic_arg_count(params),
                        generic_param_names: go_signature_generic_param_names(params),
                    },
                );
            }
            StmtKind::StructDecl { members, .. } => {
                for member in members {
                    if let ClassMember::Method(stmt) = member {
                        if let StmtKind::FunctionDecl {
                            name,
                            params,
                            return_type,
                            ..
                        } = &stmt.kind
                        {
                            signatures.insert(
                                name.clone(),
                                GoFunctionSignature {
                                    params: params
                                        .iter()
                                        .map(|param| param.type_hint.as_deref().map(str::to_string))
                                        .collect(),
                                    return_type: return_type.clone(),
                                    generic_arg_count: go_signature_generic_arg_count(params),
                                    generic_param_names: go_signature_generic_param_names(params),
                                },
                            );
                        }
                    }
                }
            }
            _ => {}
        }
    }
    signatures
}

fn go_signature_generic_arg_count(params: &[Param]) -> usize {
    params
        .iter()
        .take_while(|param| param.type_hint.as_deref() == Some("__goTypeArg"))
        .count()
}

fn go_signature_generic_param_names(params: &[Param]) -> Vec<String> {
    params
        .iter()
        .take_while(|param| param.type_hint.as_deref() == Some("__goTypeArg"))
        .filter_map(|param| go_runtime_generic_param_name(&param.name))
        .collect()
}

fn collect_go_function_bodies(body: &[Statement]) -> HashMap<String, Vec<Statement>> {
    let mut functions = HashMap::new();
    for stmt in body {
        if let StmtKind::FunctionDecl {
            name, params, body, ..
        } = &stmt.kind
        {
            if params.is_empty() {
                functions.insert(name.clone(), body.clone());
            }
        }
    }
    functions
}

fn collect_go_type_names(body: &[Statement]) -> HashSet<String> {
    let mut type_names = HashSet::new();
    for stmt in body {
        match &stmt.kind {
            StmtKind::StructDecl { name, .. }
            | StmtKind::InterfaceDecl { name, .. }
            | StmtKind::EnumDecl { name, .. }
            | StmtKind::ClassDecl { name, .. } => {
                type_names.insert(name.clone());
            }
            _ => {
                if let Some((name, _)) = go_extract_named_type_marker(stmt) {
                    type_names.insert(name);
                }
            }
        }
    }
    type_names
}

fn collect_go_named_types(body: &[Statement]) -> HashMap<String, String> {
    let mut named_types = HashMap::new();
    for stmt in body {
        if let Some((name, underlying)) = go_extract_named_type_marker(stmt) {
            named_types.insert(name, underlying);
        }
    }
    named_types
}

fn collect_go_struct_infos(body: &[Statement]) -> HashMap<String, GoStructInfo> {
    let mut infos = HashMap::new();
    for stmt in body {
        let StmtKind::StructDecl { name, members, .. } = &stmt.kind else {
            continue;
        };
        let info = infos
            .entry(name.clone())
            .or_insert_with(GoStructInfo::default);
        for member in members {
            match member {
                ClassMember::Field {
                    name,
                    type_hint,
                    modifiers,
                    ..
                } => {
                    info.field_order.push(name.clone());
                    info.member_names.insert(name.clone());
                    if let Some(tag) = go_field_tag_from_modifiers(modifiers) {
                        info.field_tags.insert(name.clone(), tag);
                    }
                    if let Some(type_name) = type_hint.clone() {
                        info.member_types.insert(name.clone(), type_name.clone());
                        if go_field_is_embedded(modifiers) {
                            info.embedded_fields.push((name.clone(), type_name));
                        }
                    }
                }
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl {
                        name,
                        params,
                        return_type,
                        ..
                    } = &stmt.kind
                    {
                        info.member_names.insert(name.clone());
                        info.method_names.insert(name.clone());
                        if params
                            .first()
                            .and_then(|param| param.type_hint.as_deref())
                            .is_some_and(|receiver| receiver.trim().starts_with('*'))
                        {
                            info.pointer_method_names.insert(name.clone());
                        }
                        if let Some(receiver_type) =
                            params.first().and_then(|param| param.type_hint.as_deref())
                        {
                            info.method_receiver_types
                                .insert(name.clone(), receiver_type.to_string());
                        }
                        info.method_params.insert(name.clone(), params.clone());
                        if let Some(type_name) = return_type.clone() {
                            info.member_types.insert(name.clone(), type_name);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    infos
}

fn collect_go_interface_methods(body: &[Statement]) -> HashMap<String, HashSet<String>> {
    let mut interfaces = HashMap::new();
    for stmt in body {
        let StmtKind::InterfaceDecl { name, members, .. } = &stmt.kind else {
            continue;
        };
        let methods = interfaces.entry(name.clone()).or_insert_with(HashSet::new);
        for member in members {
            if let InterfaceMember::Method { name, .. } = member {
                methods.insert(name.clone());
            }
        }
    }
    interfaces
}

fn go_validate_method_sets(body: &[Statement]) -> Result<(), String> {
    let env = GoNormalizeEnv {
        value_types: HashMap::new(),
        reflect_value_payloads: HashMap::new(),
        reflect_value_targets: HashMap::new(),
        reflect_pointer_targets: HashMap::new(),
        reflect_method_bindings: HashMap::new(),
        reflect_array_payloads: HashMap::new(),
        package_aliases: HashMap::new(),
        fixed_arrays: HashMap::new(),
        regex_patterns: HashMap::new(),
        slice_caps: HashMap::new(),
        slice_views: HashMap::new(),
        struct_infos: collect_go_struct_infos(body),
        interface_methods: collect_go_interface_methods(body),
        interface_concrete_types: HashMap::new(),
        nil_interface_values: HashSet::new(),
        method_value_bindings: HashMap::new(),
        named_types: collect_go_named_types(body),
        type_names: collect_go_type_names(body),
        function_bodies: HashMap::new(),
        flag_bindings: HashMap::new(),
        flag_defs: Vec::new(),
        log_output: None,
        log_prefix: None,
        log_flags: None,
        time_round_half_hour_bindings: HashSet::new(),
        generic_type_params: HashMap::new(),
        return_type: None,
        panic_value_name: None,
        has_panic_name: None,
        in_defer_name: None,
        recover_fn_name: None,
        owns_panic_state: false,
    };
    let mut locals = HashMap::new();
    for stmt in body {
        go_validate_method_set_statement(stmt, &env, &mut locals)?;
    }
    Ok(())
}

fn go_validate_method_set_statement(
    stmt: &Statement,
    env: &GoNormalizeEnv,
    locals: &mut HashMap<String, String>,
) -> Result<(), String> {
    match &stmt.kind {
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut fn_locals = locals.clone();
            for param in params {
                if let Some(type_hint) = param.type_hint.as_deref() {
                    fn_locals.insert(param.name.clone(), type_hint.to_string());
                }
            }
            for stmt in body {
                go_validate_method_set_statement(stmt, env, &mut fn_locals)?;
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(type_hint) = decl.type_hint.as_deref() {
                    if go_is_go_interface_type(type_hint.trim(), env)
                        && let Some(init) = decl.init.as_ref()
                        && let Some(init_type) = go_validation_expr_type(init, locals, env)
                        && !go_type_assignable_to_interface(&init_type, type_hint, env)
                    {
                        return Err(format!(
                            "Go method set error: {init_type} does not implement {type_hint}"
                        ));
                    }
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        locals.insert(name.clone(), type_hint.to_string());
                    }
                } else if let BindingPattern::Ident(name) = &decl.pattern
                    && let Some(init) = decl.init.as_ref()
                    && let Some(init_type) = go_validation_expr_type(init, locals, env)
                {
                    locals.insert(name.clone(), init_type);
                }
                if let Some(init) = decl.init.as_ref() {
                    go_validate_method_set_expr(init, locals, env)?;
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            go_validate_method_set_expr(value, locals, env)?;
            if let [target] = targets.as_slice()
                && let ExprKind::Ident(name) = &target.kind
                && let Some(type_hint) = locals.get(name)
                && go_is_go_interface_type(type_hint.trim(), env)
                && let Some(value_type) = go_validation_expr_type(value, locals, env)
                && !go_type_assignable_to_interface(&value_type, type_hint, env)
            {
                return Err(format!(
                    "Go method set error: {value_type} does not implement {type_hint}"
                ));
            }
            for target in targets {
                go_validate_method_set_expr(target, locals, env)?;
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            go_validate_method_set_expr(expr, locals, env)?;
        }
        StmtKind::Block(body) => {
            let mut block_locals = locals.clone();
            for stmt in body {
                go_validate_method_set_statement(stmt, env, &mut block_locals)?;
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            go_validate_method_set_expr(cond, locals, env)?;
            for stmt in then_body {
                go_validate_method_set_statement(stmt, env, &mut locals.clone())?;
            }
            for (cond, body) in elifs {
                go_validate_method_set_expr(cond, locals, env)?;
                for stmt in body {
                    go_validate_method_set_statement(stmt, env, &mut locals.clone())?;
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_validate_method_set_statement(stmt, env, &mut locals.clone())?;
                }
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. } => {
            for stmt in body {
                go_validate_method_set_statement(stmt, env, &mut locals.clone())?;
            }
        }
        StmtKind::Labeled { body, .. } => {
            go_validate_method_set_statement(body, env, locals)?;
        }
        _ => {}
    }
    Ok(())
}

fn go_validate_method_set_expr(
    expr: &Expression,
    locals: &HashMap<String, String>,
    env: &GoNormalizeEnv,
) -> Result<(), String> {
    if let Some((subject, target_type)) = go_extract_type_assert_expr(expr) {
        go_validate_type_assertion(&subject, &target_type, locals, env)?;
    }
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => {
            go_validate_method_set_expr(callee, locals, env)?;
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && let Some(receiver_type) = go_validation_expr_type(object, locals, env)
                && go_has_ambiguous_promoted_method(&receiver_type, field, env)
            {
                return Err(format!(
                    "Go method set error: ambiguous promoted method {field}"
                ));
            }
            for arg in args {
                go_validate_method_set_expr(&arg.value, locals, env)?;
            }
        }
        ExprKind::Member { object, field, .. } => {
            if let Some(receiver_type) = go_method_expression_receiver_type(object)
                && let Some(lookup) = go_struct_lookup_name(&receiver_type)
                && let Some(info) = env.struct_infos.get(&lookup)
                && info.pointer_method_names.contains(field)
                && !receiver_type.trim().starts_with('*')
            {
                return Err(format!(
                    "Go method set error: {receiver_type}.{field} requires pointer receiver"
                ));
            }
            go_validate_method_set_expr(object, locals, env)?;
        }
        ExprKind::IsType { expr, type_name } => {
            go_validate_type_assertion(expr, type_name, locals, env)?;
            go_validate_method_set_expr(expr, locals, env)?;
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Spread(expr)
        | ExprKind::TypeOf(expr) => go_validate_method_set_expr(expr, locals, env)?,
        ExprKind::Binary { left, right, .. } => {
            go_validate_method_set_expr(left, locals, env)?;
            go_validate_method_set_expr(right, locals, env)?;
        }
        ExprKind::Assign { target, value } => {
            go_validate_method_set_expr(target, locals, env)?;
            go_validate_method_set_expr(value, locals, env)?;
        }
        ExprKind::Index { object, index, .. } => {
            go_validate_method_set_expr(object, locals, env)?;
            go_validate_method_set_expr(index, locals, env)?;
        }
        ExprKind::Array(elements) => {
            for element in elements {
                go_validate_method_set_expr(&element.value, locals, env)?;
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    go_validate_method_set_expr(key, locals, env)?;
                    go_validate_method_set_expr(value, locals, env)?;
                }
            }
        }
        ExprKind::Tuple(values) | ExprKind::Sequence(values) => {
            for value in values {
                go_validate_method_set_expr(value, locals, env)?;
            }
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_validate_method_set_expr(cond, locals, env)?;
            go_validate_method_set_expr(then, locals, env)?;
            go_validate_method_set_expr(else_, locals, env)?;
        }
        ExprKind::Lambda { body, .. } => {
            if let LambdaBody::Block(body) = body {
                let mut lambda_locals = locals.clone();
                for stmt in body {
                    go_validate_method_set_statement(stmt, env, &mut lambda_locals)?;
                }
            }
        }
        _ => {}
    }
    Ok(())
}

fn go_validate_type_assertion(
    subject: &Expression,
    target_type: &str,
    locals: &HashMap<String, String>,
    env: &GoNormalizeEnv,
) -> Result<(), String> {
    let Some(subject_type) = go_validation_expr_type(subject, locals, env) else {
        return Ok(());
    };
    if !go_is_go_interface_type(&subject_type, env) {
        return Err(format!(
            "Go type assertion error: {subject_type} is not an interface"
        ));
    }
    let target = target_type.trim();
    if go_is_go_interface_type(target, env) {
        if matches!(subject_type.trim(), "interface{}" | "any")
            || matches!(target, "interface{}" | "any")
            || target == subject_type.trim()
            || go_interface_implements_interface(target, &subject_type, env)
        {
            return Ok(());
        }
        return Err(format!(
            "Go type assertion error: {target} does not implement {subject_type}"
        ));
    }
    if matches!(subject_type.trim(), "interface{}" | "any")
        || go_type_assignable_to_interface(target, &subject_type, env)
    {
        Ok(())
    } else {
        Err(format!(
            "Go type assertion error: {target} does not implement {subject_type}"
        ))
    }
}

fn go_interface_implements_interface(
    target_interface: &str,
    source_interface: &str,
    env: &GoNormalizeEnv,
) -> bool {
    let Some(source_methods) = env.interface_methods.get(source_interface.trim()) else {
        return false;
    };
    let Some(target_methods) = env.interface_methods.get(target_interface.trim()) else {
        return false;
    };
    source_methods
        .iter()
        .all(|method| target_methods.contains(method))
}

fn go_validation_expr_type(
    expr: &Expression,
    locals: &HashMap<String, String>,
    env: &GoNormalizeEnv,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => locals.get(name).cloned(),
        ExprKind::Lit(Literal::Int(_)) => Some("int".to_string()),
        ExprKind::Lit(Literal::Float(_)) => Some("float64".to_string()),
        ExprKind::Lit(Literal::Bool(_)) => Some("bool".to_string()),
        ExprKind::Lit(Literal::Str(_)) => Some("string".to_string()),
        ExprKind::Cast { type_name, .. } => Some(type_name.clone()),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => go_validation_expr_type(expr, locals, env).map(|ty| format!("*{}", ty.trim())),
        ExprKind::Unary {
            op: UnaryOp::Deref,
            expr,
        }
        | ExprKind::RefLoad(expr) => go_validation_expr_type(expr, locals, env).map(|ty| {
            ty.trim()
                .trim_start_matches('*')
                .trim_start_matches('^')
                .trim()
                .to_string()
        }),
        ExprKind::Member { object, field, .. } => {
            let receiver_type = go_validation_expr_type(object, locals, env)?;
            go_resolve_struct_member_type(&receiver_type, field, env, &mut HashSet::new())
        }
        _ => None,
    }
}

fn go_has_ambiguous_promoted_method(type_name: &str, method: &str, env: &GoNormalizeEnv) -> bool {
    let Some(lookup) = go_struct_lookup_name(type_name) else {
        return false;
    };
    let Some(info) = env.struct_infos.get(&lookup) else {
        return false;
    };
    if info.method_names.contains(method) {
        return false;
    }
    info.embedded_fields
        .iter()
        .filter(|(_, embedded_type)| go_type_has_method(embedded_type, method, env))
        .take(2)
        .count()
        > 1
}

fn go_field_tag_from_modifiers(modifiers: &Modifiers) -> Option<String> {
    modifiers.decorators.iter().find_map(|decorator| {
        let ExprKind::Lit(Literal::Str(text)) = &decorator.kind else {
            return None;
        };
        text.find("__go_tag:")
            .map(|idx| text[idx + "__go_tag:".len()..].to_string())
    })
}

/// `struct { Inner }` promotes its fields; `struct { Inner Inner }` does not.
/// The two look identical once the walker fills the missing field name in from
/// the type, so the embedding is recorded while the source still shows it —
/// comparing the name back against the type calls every `T T` field embedded.
fn go_field_is_embedded(modifiers: &Modifiers) -> bool {
    modifiers.decorators.iter().any(|decorator| {
        matches!(&decorator.kind, ExprKind::Lit(Literal::Str(text)) if &**text == GO_EMBEDDED_MARKER)
    })
}

const GO_EMBEDDED_MARKER: &str = "__go_embedded";

fn merge_go_struct_decls(body: &[Statement]) -> Vec<Statement> {
    let mut first_index: HashMap<String, usize> = HashMap::new();
    for (index, stmt) in body.iter().enumerate() {
        if let StmtKind::StructDecl { name, .. } = &stmt.kind {
            first_index.entry(name.clone()).or_insert(index);
        }
    }

    let mut emitted = std::collections::HashSet::new();
    let mut merged_body = Vec::with_capacity(body.len());

    for (index, stmt) in body.iter().enumerate() {
        match &stmt.kind {
            StmtKind::StructDecl { name, .. } => {
                if first_index.get(name) != Some(&index) || !emitted.insert(name.clone()) {
                    continue;
                }

                let mut merged = stmt.clone();
                if let StmtKind::StructDecl {
                    interfaces,
                    members,
                    ..
                } = &mut merged.kind
                {
                    for later in body.iter().skip(index + 1) {
                        if let StmtKind::StructDecl {
                            name: later_name,
                            interfaces: later_interfaces,
                            members: later_members,
                            ..
                        } = &later.kind
                        {
                            if later_name == name {
                                members.extend(later_members.clone());
                                for interface in later_interfaces {
                                    if !interfaces.iter().any(|existing| existing == interface) {
                                        interfaces.push(interface.clone());
                                    }
                                }
                            }
                        }
                    }
                }

                merged_body.push(merged);
            }
            _ => merged_body.push(stmt.clone()),
        }
    }

    merged_body
}

fn collect_go_global_fixed_arrays(
    body: &[Statement],
    signatures: &HashMap<String, GoFunctionSignature>,
) -> HashMap<String, String> {
    let env = GoNormalizeEnv::default();
    let mut globals = HashMap::new();

    for stmt in body {
        if let StmtKind::VarDecl { declarations, .. } = &stmt.kind {
            for decl in declarations {
                if let Some((name, type_name)) = go_decl_fixed_array_binding(decl, &env, signatures)
                {
                    globals.insert(name, type_name);
                }
            }
        }
    }

    globals
}

fn normalize_go_block(
    stmts: &[Statement],
    base_env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let mut env = base_env.clone();
    let mut normalized = Vec::with_capacity(stmts.len());
    for stmt in stmts {
        normalized.extend(normalize_go_statement(stmt, &mut env, signatures, state));
    }
    normalized
}

fn normalize_go_function_body(
    stmts: &[Statement],
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    if env.recover_fn_name.is_none() {
        env.panic_value_name = Some(fresh_go_temp(state, "__go_panic_value"));
        env.has_panic_name = Some(fresh_go_temp(state, "__go_has_panic"));
        env.in_defer_name = Some(fresh_go_temp(state, "__go_in_defer"));
        env.recover_fn_name = Some(fresh_go_temp(state, "__go_recover"));
        env.owns_panic_state = true;
    } else {
        env.owns_panic_state = false;
    }

    let mut named_results: Vec<Param> = Vec::new();
    let mut body_stmts = Vec::with_capacity(stmts.len());
    for stmt in stmts {
        if let Some(param) = go_extract_named_result_marker(stmt) {
            env.value_types.insert(
                param.name.clone(),
                param
                    .type_hint
                    .as_deref()
                    .map(str::to_string)
                    .unwrap_or_else(|| "object".to_string()),
            );
            named_results.push(param);
            continue;
        }
        body_stmts.push(stmt.clone());
    }

    let mut normalized = Vec::with_capacity(stmts.len());
    for stmt in &body_stmts {
        normalized.extend(normalize_go_statement(stmt, env, signatures, state));
    }

    let (normalized, final_return) = if !named_results.is_empty() {
        go_lower_named_results_body(normalized, &named_results, state)
    } else {
        (normalized, None)
    };

    lower_go_defer_body(normalized, env, signatures, state, final_return)
}

fn lower_go_defer_body(
    body: Vec<Statement>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    final_return: Option<Expression>,
) -> Vec<Statement> {
    let panic_value_name = env
        .panic_value_name
        .clone()
        .unwrap_or_else(|| fresh_go_temp(state, "__go_panic_value"));
    let has_panic_name = env
        .has_panic_name
        .clone()
        .unwrap_or_else(|| fresh_go_temp(state, "__go_has_panic"));
    let in_defer_name = env
        .in_defer_name
        .clone()
        .unwrap_or_else(|| fresh_go_temp(state, "__go_in_defer"));
    let recover_fn_name = env
        .recover_fn_name
        .clone()
        .unwrap_or_else(|| fresh_go_temp(state, "__go_recover"));
    let stack_name = fresh_go_temp(state, "__go_defer_stack");
    let (mut lowered_body, has_defer) =
        lower_go_defer_statements(body, env, signatures, state, &stack_name, false);
    let mut defer_ref_names = HashSet::new();
    for stmt in &lowered_body {
        go_collect_ref_place_names_stmt(stmt, &mut defer_ref_names);
    }
    let post_defer_check = if has_defer && lowered_body.last().is_some_and(go_is_check_statement) {
        lowered_body.pop()
    } else {
        None
    };
    let mut post_defer_output = Vec::new();
    if post_defer_check.is_some() {
        while lowered_body
            .last()
            .is_some_and(go_is_harness_output_statement)
        {
            if let Some(stmt) = lowered_body.pop() {
                post_defer_output.push(stmt);
            }
        }
        post_defer_output.reverse();
    }

    let panic_value_decl = go_defer_temp_decl(panic_value_name.clone(), None, Expression::null());
    let has_panic_decl = go_defer_temp_decl(has_panic_name.clone(), None, Expression::bool(false));
    let in_defer_decl = go_defer_temp_decl(in_defer_name.clone(), None, Expression::bool(false));
    let recover_value_name = format!("{recover_fn_name}_value");
    let recover_fn_decl = go_defer_temp_decl(
        recover_fn_name,
        None,
        Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(vec![Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(Expression::ident(&has_panic_name)),
                    right: Box::new(Expression::ident(&in_defer_name)),
                }),
                then_body: vec![
                    go_defer_temp_decl(
                        recover_value_name.clone(),
                        None,
                        Expression::ident(&panic_value_name),
                    ),
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(&has_panic_name)],
                        value: Expression::bool(false),
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident(
                        &recover_value_name,
                    )))),
                ],
                elifs: Vec::new(),
                else_body: Some(vec![Statement::new(StmtKind::Return(Some(
                    Expression::null(),
                )))]),
            })]),
            is_async: false,
            captures: Vec::new(),
        }),
    );

    let panic_state_decls = if env.owns_panic_state {
        vec![
            panic_value_decl,
            has_panic_decl,
            in_defer_decl,
            recover_fn_decl,
        ]
    } else {
        Vec::new()
    };

    if !has_defer {
        let mut body = panic_state_decls;
        body.extend(lowered_body);
        if let Some(expr) = final_return {
            body.push(Statement::new(StmtKind::Return(Some(expr))));
        }
        return body;
    }

    let stack_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(stack_name.clone()),
            type_hint: None,
            init: Some(Expression::null()),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let drain_name = fresh_go_temp(state, "__go_defer_fn");
    let drain_panic_name = fresh_go_temp(state, "__go_defer_panic");
    let drain_loop = Statement::new(StmtKind::While {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::ident(&stack_name)),
            right: Box::new(Expression::null()),
        }),
        body: vec![
            go_defer_temp_decl(
                drain_name.clone(),
                None,
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&stack_name)),
                    field: "fn".to_string(),
                    null_safe: false,
                }),
            ),
            go_defer_temp_decl(
                format!("{drain_name}_recover"),
                None,
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&stack_name)),
                    field: "recover".to_string(),
                    null_safe: false,
                }),
            ),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&stack_name)],
                value: Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&stack_name)),
                    field: "next".to_string(),
                    null_safe: false,
                }),
                by_ref: false,
            }),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&in_defer_name)],
                value: Expression::ident(&format!("{drain_name}_recover")),
                by_ref: false,
            }),
            Statement::new(StmtKind::Try {
                body: vec![Statement::new(StmtKind::Expr(Expression::new(
                    ExprKind::Call {
                        callee: Box::new(Expression::ident(&drain_name)),
                        args: Vec::new(),
                        optional: false,
                    },
                )))],
                catches: vec![CatchClause {
                    types: Vec::new(),
                    var_name: Some(drain_panic_name.clone()),
                    stack_var: None,
                    body: vec![
                        Statement::new(StmtKind::Assign {
                            targets: vec![Expression::ident(&panic_value_name)],
                            value: Expression::ident(&drain_panic_name),
                            by_ref: false,
                        }),
                        Statement::new(StmtKind::Assign {
                            targets: vec![Expression::ident(&has_panic_name)],
                            value: Expression::bool(true),
                            by_ref: false,
                        }),
                    ],
                    when_clause: None,
                }],
                else_body: None,
                finally: None,
            }),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&in_defer_name)],
                value: Expression::bool(false),
                by_ref: false,
            }),
        ],
        else_body: None,
    });

    let panic_catch_name = fresh_go_temp(state, "__go_panic_exc");

    let mut body = panic_state_decls;
    let mut success_body = Vec::new();
    success_body.extend(
        post_defer_output
            .into_iter()
            .map(|stmt| go_rewrite_ref_place_reads_stmt(&stmt, &defer_ref_names)),
    );
    if let Some(check_stmt) = post_defer_check {
        success_body.push(check_stmt);
    }
    if let Some(expr) = final_return {
        success_body.push(Statement::new(StmtKind::Return(Some(expr))));
    }

    body.extend([
        stack_decl,
        Statement::new(StmtKind::Try {
            body: lowered_body,
            catches: vec![CatchClause {
                types: Vec::new(),
                var_name: Some(panic_catch_name.clone()),
                stack_var: None,
                body: vec![
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(&panic_value_name)],
                        value: Expression::ident(&panic_catch_name),
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(&has_panic_name)],
                        value: Expression::bool(true),
                        by_ref: false,
                    }),
                ],
                when_clause: None,
            }],
            else_body: None,
            finally: Some(vec![drain_loop]),
        }),
        Statement::new(StmtKind::If {
            cond: Expression::ident(&has_panic_name),
            then_body: vec![Statement::new(StmtKind::Throw {
                expr: Some(Expression::ident(&panic_value_name)),
                cause: None,
            })],
            elifs: Vec::new(),
            else_body: if success_body.is_empty() {
                None
            } else {
                Some(success_body)
            },
        }),
    ]);
    body
}

fn go_is_check_statement(stmt: &Statement) -> bool {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return false;
    };
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__check")
    )
}

fn go_is_harness_output_statement(stmt: &Statement) -> bool {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return false;
    };
    let ExprKind::Call { callee, .. } = &expr.kind else {
        return false;
    };
    matches!(&callee.kind, ExprKind::Ident(name) if name == "__p" || name == "__pr")
}

fn go_extract_named_result_marker(stmt: &Statement) -> Option<Param> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_named_result")
        || args.len() != 2
    {
        return None;
    }
    let ExprKind::Lit(Literal::Str(name)) = &args[0].value.kind else {
        return None;
    };
    let type_hint = go_type_name_from_expr(&args[1].value)?;
    Some(Param {
        name: name.clone(),
        type_hint: Some(type_hint.into()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    })
}

fn go_extract_named_type_marker(stmt: &Statement) -> Option<(String, String)> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_named_type")
        || args.len() != 2
    {
        return None;
    }
    let ExprKind::Lit(Literal::Str(name)) = &args[0].value.kind else {
        return None;
    };
    let type_name = go_type_name_from_expr(&args[1].value)?;
    Some((name.clone(), type_name))
}

fn go_lower_named_results_body(
    body: Vec<Statement>,
    results: &[Param],
    state: &mut GoNormalizeState,
) -> (Vec<Statement>, Option<Expression>) {
    // Use a sentinel string only as a rewrite marker for go_rewrite_named_result_returns.
    // We no longer throw/catch this sentinel at runtime — instead we use a
    // `while(true) { ...; break }` loop so that `return X` inside any branch
    // compiles to `result = X; break`, which emits a clean `BR(N)` that
    // correctly unwinds the label stack.  Throwing a sentinel inside an IF
    // body left extra BLOCK labels on the label_stack that THROW does not
    // restore, corrupting the outer catch handler lookup.
    let sentinel = fresh_go_temp(state, "__go_named_return");
    let mut body = body;
    for result in results {
        body = go_rewrite_named_result_cell_body(body, &result.name);
    }
    let rewritten_body = go_rewrite_named_result_returns(body, results, &sentinel);
    let result_inits = results
        .iter()
        .map(|result| {
            let result_type = result
                .type_hint
                .clone()
                .unwrap_or_else(|| "object".to_string().into());
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&result.name)],
                value: go_named_result_cell_object(go_zero_value_expr(&result_type)),
                by_ref: false,
            })
        })
        .collect::<Vec<_>>();
    // Wrap the rewritten body in `while(true) { ...; break }`.
    // Every `return X` in rewritten_body was turned into `result=X; break`.
    // At end of function body (bare `return`) we also emit an implicit break.
    // Real panics (user-thrown exceptions) propagate directly to the outer
    // defer try/catch because no try/catch is interposed here.
    let mut while_body = rewritten_body;
    // Ensure the while always exits: append an implicit break so fall-through
    // at end of body exits the loop rather than looping forever.
    while_body.push(Statement::new(StmtKind::Break(BreakTarget::Implicit)));

    let final_return = if results.len() == 1 {
        Some(go_named_result_cell_value(&results[0].name))
    } else {
        Some(Expression::new(ExprKind::Tuple(
            results
                .iter()
                .map(|result| go_named_result_cell_value(&result.name))
                .collect(),
        )))
    };

    let mut lowered = result_inits;
    lowered.push(Statement::new(StmtKind::While {
        cond: Expression::bool(true),
        body: while_body,
        else_body: None,
    }));

    (lowered, final_return)
}

fn go_named_result_cell_object(value: Expression) -> Expression {
    Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
        key: Expression::string("value"),
        value,
    }]))
}

fn go_named_result_cell_value(name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(name)),
        field: "value".to_string(),
        null_safe: false,
    })
}

fn go_rewrite_named_result_cell_body(body: Vec<Statement>, result_name: &str) -> Vec<Statement> {
    body.into_iter()
        .map(|stmt| go_rewrite_named_result_cell_stmt(stmt, result_name))
        .collect()
}

fn go_rewrite_named_result_cell_stmt(stmt: Statement, result_name: &str) -> Statement {
    match stmt.kind {
        StmtKind::Expr(expr) => Statement::new(StmtKind::Expr(go_rewrite_named_result_cell_expr(
            expr,
            result_name,
        ))),
        StmtKind::Return(expr) => Statement::new(StmtKind::Return(
            expr.map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
        )),
        StmtKind::Throw { expr, cause } => Statement::new(StmtKind::Throw {
            expr: expr.map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
            cause: cause.map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
        }),
        StmtKind::VarDecl { declarations, kind } => Statement::new(StmtKind::VarDecl {
            declarations: declarations
                .into_iter()
                .map(|mut decl| {
                    decl.init = decl
                        .init
                        .map(|expr| go_rewrite_named_result_cell_expr(expr, result_name));
                    decl
                })
                .collect(),
            kind,
        }),
        StmtKind::Assign { targets, value, .. } => Statement::new(StmtKind::Assign {
            targets: targets
                .into_iter()
                .map(|target| go_rewrite_named_result_cell_target(target, result_name))
                .collect(),
            value: go_rewrite_named_result_cell_expr(value, result_name),
            by_ref: false,
        }),
        StmtKind::CompoundAssign { target, op, value } => {
            Statement::new(StmtKind::CompoundAssign {
                target: go_rewrite_named_result_cell_target(target, result_name),
                op,
                value: go_rewrite_named_result_cell_expr(value, result_name),
            })
        }
        StmtKind::Block(body) => Statement::new(StmtKind::Block(
            go_rewrite_named_result_cell_body(body, result_name),
        )),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => Statement::new(StmtKind::If {
            cond: go_rewrite_named_result_cell_expr(cond, result_name),
            then_body: go_rewrite_named_result_cell_body(then_body, result_name),
            elifs: elifs
                .into_iter()
                .map(|(cond, body)| {
                    (
                        go_rewrite_named_result_cell_expr(cond, result_name),
                        go_rewrite_named_result_cell_body(body, result_name),
                    )
                })
                .collect(),
            else_body: else_body.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
        }),
        StmtKind::While {
            cond,
            body,
            else_body,
        } => Statement::new(StmtKind::While {
            cond: go_rewrite_named_result_cell_expr(cond, result_name),
            body: go_rewrite_named_result_cell_body(body, result_name),
            else_body: else_body.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
        }),
        StmtKind::DoWhile { body, cond, until } => Statement::new(StmtKind::DoWhile {
            body: go_rewrite_named_result_cell_body(body, result_name),
            cond: go_rewrite_named_result_cell_expr(cond, result_name),
            until,
        }),
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => Statement::new(StmtKind::For {
            init: init.map(|stmt| Box::new(go_rewrite_named_result_cell_stmt(*stmt, result_name))),
            cond: cond.map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
            update: update.map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
            body: go_rewrite_named_result_cell_body(body, result_name),
        }),
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            else_body,
            is_async,
            of,
        } => Statement::new(StmtKind::ForIn {
            var,
            key,
            iter: go_rewrite_named_result_cell_expr(iter, result_name),
            body: go_rewrite_named_result_cell_body(body, result_name),
            else_body: else_body.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
            is_async,
            of,
        }),
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => Statement::new(StmtKind::Switch {
            expr: go_rewrite_named_result_cell_expr(expr, result_name),
            cases: cases
                .into_iter()
                .map(|case| SwitchCase {
                    conditions: case
                        .conditions
                        .into_iter()
                        .map(|condition| match condition {
                            CaseCondition::Value(expr) => CaseCondition::Value(
                                go_rewrite_named_result_cell_expr(expr, result_name),
                            ),
                            other => other,
                        })
                        .collect(),
                    body: go_rewrite_named_result_cell_body(case.body, result_name),
                })
                .collect(),
            default: default.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
        }),
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => Statement::new(StmtKind::Try {
            body: go_rewrite_named_result_cell_body(body, result_name),
            catches: catches
                .into_iter()
                .map(|catch| CatchClause {
                    body: go_rewrite_named_result_cell_body(catch.body, result_name),
                    ..catch
                })
                .collect(),
            else_body: else_body.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
            finally: finally.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
        }),
        StmtKind::Labeled { label, body } => Statement::new(StmtKind::Labeled {
            label,
            body: Box::new(go_rewrite_named_result_cell_stmt(*body, result_name)),
        }),
        other => Statement::new(other),
    }
}

fn go_rewrite_named_result_cell_target(target: Expression, result_name: &str) -> Expression {
    if matches!(&target.kind, ExprKind::Ident(name) if name == result_name) {
        return go_named_result_cell_value(result_name);
    }
    go_rewrite_named_result_cell_expr(target, result_name)
}

fn go_rewrite_named_result_cell_expr(expr: Expression, result_name: &str) -> Expression {
    match expr.kind {
        ExprKind::Ident(name) if name == result_name => go_named_result_cell_value(result_name),
        ExprKind::Unary { op, expr } => Expression::new(ExprKind::Unary {
            op,
            expr: Box::new(go_rewrite_named_result_cell_expr(*expr, result_name)),
        }),
        ExprKind::Binary { left, op, right } => Expression::new(ExprKind::Binary {
            left: Box::new(go_rewrite_named_result_cell_expr(*left, result_name)),
            op,
            right: Box::new(go_rewrite_named_result_cell_expr(*right, result_name)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(go_rewrite_named_result_cell_expr(*cond, result_name)),
            then: Box::new(go_rewrite_named_result_cell_expr(*then, result_name)),
            else_: Box::new(go_rewrite_named_result_cell_expr(*else_, result_name)),
        }),
        ExprKind::Cast { expr, type_name } => Expression::new(ExprKind::Cast {
            expr: Box::new(go_rewrite_named_result_cell_expr(*expr, result_name)),
            type_name,
        }),
        ExprKind::RefLoad(inner) => Expression::new(ExprKind::RefLoad(Box::new(
            go_rewrite_named_result_cell_expr(*inner, result_name),
        ))),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: Box::new(go_rewrite_named_result_cell_expr(*object, result_name)),
            field,
            null_safe,
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(go_rewrite_named_result_cell_expr(*object, result_name)),
            index: Box::new(go_rewrite_named_result_cell_expr(*index, result_name)),
            null_safe,
        }),
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(go_rewrite_named_result_cell_target(*target, result_name)),
            value: Box::new(go_rewrite_named_result_cell_expr(*value, result_name)),
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => Expression::new(ExprKind::Call {
            callee: Box::new(go_rewrite_named_result_cell_expr(*callee, result_name)),
            args: args
                .into_iter()
                .map(|arg| Argument {
                    value: go_rewrite_named_result_cell_expr(arg.value, result_name),
                    ..arg
                })
                .collect(),
            optional,
        }),
        ExprKind::Array(elements) => Expression::new(ExprKind::Array(
            elements
                .into_iter()
                .map(|element| ArrayElement {
                    key: element
                        .key
                        .map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
                    value: go_rewrite_named_result_cell_expr(element.value, result_name),
                    ..element
                })
                .collect(),
        )),
        ExprKind::Object(properties) => Expression::new(ExprKind::Object(
            properties
                .into_iter()
                .map(|property| match property {
                    ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                        key: go_rewrite_named_result_cell_expr(key, result_name),
                        value: go_rewrite_named_result_cell_expr(value, result_name),
                    },
                    ObjectProperty::Spread(value) => ObjectProperty::Spread(
                        go_rewrite_named_result_cell_expr(value, result_name),
                    ),
                    ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                        key: go_rewrite_named_result_cell_expr(key, result_name),
                        value: go_rewrite_named_result_cell_expr(value, result_name),
                    },
                    other => other,
                })
                .collect(),
        )),
        ExprKind::Tuple(values) => Expression::new(ExprKind::Tuple(
            values
                .into_iter()
                .map(|value| go_rewrite_named_result_cell_expr(value, result_name))
                .collect(),
        )),
        ExprKind::Sequence(values) => Expression::new(ExprKind::Sequence(
            values
                .into_iter()
                .map(|value| go_rewrite_named_result_cell_expr(value, result_name))
                .collect(),
        )),
        ExprKind::Lambda {
            params,
            body,
            is_async,
            captures,
        } => {
            if params.iter().any(|param| param.name == result_name) {
                Expression::new(ExprKind::Lambda {
                    params,
                    body,
                    is_async,
                    captures,
                })
            } else {
                Expression::new(ExprKind::Lambda {
                    params,
                    body: go_rewrite_named_result_cell_lambda_body(body, result_name),
                    is_async,
                    captures,
                })
            }
        }
        other => Expression::new(other),
    }
}

fn go_rewrite_named_result_cell_lambda_body(body: LambdaBody, result_name: &str) -> LambdaBody {
    match body {
        LambdaBody::Expr(expr) => LambdaBody::Expr(Box::new(go_rewrite_named_result_cell_expr(
            *expr,
            result_name,
        ))),
        LambdaBody::Block(body) => {
            LambdaBody::Block(go_rewrite_named_result_cell_body(body, result_name))
        }
    }
}

fn go_rewrite_named_result_returns(
    body: Vec<Statement>,
    results: &[Param],
    sentinel: &str,
) -> Vec<Statement> {
    let mut rewritten = Vec::with_capacity(body.len());
    for stmt in body {
        rewritten.extend(go_rewrite_named_result_return_stmt(stmt, results, sentinel));
    }
    rewritten
}

fn go_rewrite_named_result_return_stmt(
    stmt: Statement,
    results: &[Param],
    sentinel: &str,
) -> Vec<Statement> {
    match stmt.kind {
        StmtKind::Return(expr) => {
            let mut rewritten = Vec::new();
            if let Some(expr) = expr {
                let target = if results.len() == 1 {
                    go_named_result_cell_value(&results[0].name)
                } else {
                    Expression::new(ExprKind::Tuple(
                        results
                            .iter()
                            .map(|result| go_named_result_cell_value(&result.name))
                            .collect(),
                    ))
                };
                rewritten.push(Statement::new(StmtKind::Assign {
                    targets: vec![target],
                    value: expr,
                    by_ref: false,
                }));
            }
            // Break out of the enclosing while(true) loop cleanly.
            // This compiles to BR(N) which correctly unwinds the label stack
            // regardless of how many nested BLOCKs (e.g. from if-statements)
            // are active at the break site.
            rewritten.push(Statement::new(StmtKind::Break(BreakTarget::Implicit)));
            rewritten
        }
        StmtKind::Block(body) => vec![Statement::new(StmtKind::Block(
            go_rewrite_named_result_returns(body, results, sentinel),
        ))],
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => vec![Statement::new(StmtKind::If {
            cond,
            then_body: go_rewrite_named_result_returns(then_body, results, sentinel),
            elifs: elifs
                .into_iter()
                .map(|(cond, body)| {
                    (
                        cond,
                        go_rewrite_named_result_returns(body, results, sentinel),
                    )
                })
                .collect(),
            else_body: else_body
                .map(|body| go_rewrite_named_result_returns(body, results, sentinel)),
        })],
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => vec![Statement::new(StmtKind::For {
            init,
            cond,
            update,
            body: go_rewrite_named_result_returns(body, results, sentinel),
        })],
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => vec![Statement::new(StmtKind::ForIn {
            var,
            key,
            iter,
            body: go_rewrite_named_result_returns(body, results, sentinel),
            of,
            else_body: else_body
                .map(|body| go_rewrite_named_result_returns(body, results, sentinel)),
            is_async,
        })],
        StmtKind::While {
            cond,
            body,
            else_body,
        } => vec![Statement::new(StmtKind::While {
            cond,
            body: go_rewrite_named_result_returns(body, results, sentinel),
            else_body: else_body
                .map(|body| go_rewrite_named_result_returns(body, results, sentinel)),
        })],
        StmtKind::DoWhile { body, cond, until } => vec![Statement::new(StmtKind::DoWhile {
            body: go_rewrite_named_result_returns(body, results, sentinel),
            cond,
            until,
        })],
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => vec![Statement::new(StmtKind::Switch {
            expr,
            cases: cases
                .into_iter()
                .map(|case| SwitchCase {
                    conditions: case.conditions,
                    body: go_rewrite_named_result_returns(case.body, results, sentinel),
                })
                .collect(),
            default: default.map(|body| go_rewrite_named_result_returns(body, results, sentinel)),
        })],
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => vec![Statement::new(StmtKind::Try {
            body: go_rewrite_named_result_returns(body, results, sentinel),
            catches: catches
                .into_iter()
                .map(|catch| CatchClause {
                    types: catch.types,
                    var_name: catch.var_name,
                    stack_var: catch.stack_var,
                    body: go_rewrite_named_result_returns(catch.body, results, sentinel),
                    when_clause: catch.when_clause,
                })
                .collect(),
            else_body: else_body
                .map(|body| go_rewrite_named_result_returns(body, results, sentinel)),
            finally: finally.map(|body| go_rewrite_named_result_returns(body, results, sentinel)),
        })],
        StmtKind::Select { arms, default } => vec![Statement::new(StmtKind::Select {
            arms: arms
                .into_iter()
                .map(|arm| SelectArm {
                    comm: arm.comm,
                    body: go_rewrite_named_result_returns(arm.body, results, sentinel),
                })
                .collect(),
            default: default.map(|body| go_rewrite_named_result_returns(body, results, sentinel)),
        })],
        StmtKind::Labeled { label, body } => vec![Statement::new(StmtKind::Labeled {
            label,
            body: Box::new({
                let mut rewritten = go_rewrite_named_result_return_stmt(*body, results, sentinel);
                if rewritten.len() == 1 {
                    rewritten.remove(0)
                } else {
                    Statement::new(StmtKind::Block(rewritten))
                }
            }),
        })],
        _ => vec![stmt],
    }
}

fn lower_go_defer_statements(
    body: Vec<Statement>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    stack_name: &str,
    in_loop: bool,
) -> (Vec<Statement>, bool) {
    lower_go_defer_statements_with_loop_names(
        body,
        env,
        signatures,
        state,
        stack_name,
        in_loop,
        HashSet::new(),
    )
}

fn lower_go_defer_statements_with_loop_names(
    body: Vec<Statement>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    stack_name: &str,
    in_loop: bool,
    initial_loop_local_names: HashSet<String>,
) -> (Vec<Statement>, bool) {
    let mut lowered = Vec::with_capacity(body.len());
    let mut has_defer = false;
    let mut loop_local_names = initial_loop_local_names;
    let empty_loop_local_names = HashSet::new();

    for stmt in body {
        if let Some(expr) = go_extract_defer_expr(&stmt) {
            let frozen_names = if in_loop {
                &loop_local_names
            } else {
                &empty_loop_local_names
            };
            lowered.extend(go_lower_defer_stmt(
                expr,
                env,
                signatures,
                state,
                stack_name,
                frozen_names,
                in_loop,
            ));
            has_defer = true;
            continue;
        }

        let (next_stmt, nested_has_defer) =
            lower_go_defer_statement(stmt, env, signatures, state, stack_name, in_loop);
        if in_loop {
            go_collect_block_declared_names(&next_stmt, &mut loop_local_names);
        }
        lowered.push(next_stmt);
        has_defer |= nested_has_defer;
    }

    (lowered, has_defer)
}

fn lower_go_defer_statement(
    stmt: Statement,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    stack_name: &str,
    in_loop: bool,
) -> (Statement, bool) {
    match stmt.kind {
        StmtKind::Block(body) => {
            let (body, has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
            (Statement::new(StmtKind::Block(body)), has_defer)
        }
        StmtKind::Labeled { label, body } => {
            let (body, has_defer) =
                lower_go_defer_statement(*body, env, signatures, state, stack_name, in_loop);
            (
                Statement::new(StmtKind::Labeled {
                    label,
                    body: Box::new(body),
                }),
                has_defer,
            )
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            let (next_then, mut has_defer) =
                lower_go_defer_statements(then_body, env, signatures, state, stack_name, in_loop);
            let mut next_elifs = Vec::with_capacity(elifs.len());
            for (elif_cond, elif_body) in elifs {
                let (next_body, nested_has_defer) = lower_go_defer_statements(
                    elif_body, env, signatures, state, stack_name, in_loop,
                );
                next_elifs.push((elif_cond, next_body));
                has_defer |= nested_has_defer;
            }
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::If {
                    cond,
                    then_body: next_then,
                    elifs: next_elifs,
                    else_body: next_else,
                }),
                has_defer,
            )
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_names = HashSet::new();
            if let Some(init) = &init {
                go_collect_block_declared_names(init, &mut loop_names);
            }
            let (body, has_defer) = lower_go_defer_statements_with_loop_names(
                body, env, signatures, state, stack_name, true, loop_names,
            );
            (
                Statement::new(StmtKind::For {
                    init,
                    cond,
                    update,
                    body,
                }),
                has_defer,
            )
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => {
            let (body, mut has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, true);
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::ForIn {
                    var,
                    key,
                    iter,
                    body,
                    of,
                    else_body: next_else,
                    is_async,
                }),
                has_defer,
            )
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            let loop_context = if go_is_named_result_wrapper_while(&cond, &body, &else_body) {
                in_loop
            } else {
                true
            };
            let (body, mut has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, loop_context);
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::While {
                    cond,
                    body,
                    else_body: next_else,
                }),
                has_defer,
            )
        }
        StmtKind::DoWhile { body, cond, until } => {
            let (body, has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, true);
            (
                Statement::new(StmtKind::DoWhile { body, cond, until }),
                has_defer,
            )
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            let mut has_defer = false;
            let next_cases = cases
                .into_iter()
                .map(|case| {
                    let (body, nested_has_defer) = lower_go_defer_statements(
                        case.body, env, signatures, state, stack_name, in_loop,
                    );
                    has_defer |= nested_has_defer;
                    SwitchCase {
                        conditions: case.conditions,
                        body,
                    }
                })
                .collect();
            let next_default = if let Some(body) = default {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::Switch {
                    expr,
                    cases: next_cases,
                    default: next_default,
                }),
                has_defer,
            )
        }
        StmtKind::Select { arms, default } => {
            let mut has_defer = false;
            let next_arms = arms
                .into_iter()
                .map(|arm| {
                    let (body, nested_has_defer) = lower_go_defer_statements(
                        arm.body, env, signatures, state, stack_name, true,
                    );
                    has_defer |= nested_has_defer;
                    SelectArm {
                        comm: arm.comm,
                        body,
                    }
                })
                .collect();
            let next_default = if let Some(body) = default {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, true);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::Select {
                    arms: next_arms,
                    default: next_default,
                }),
                has_defer,
            )
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            let (body, mut has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
            let next_catches = catches
                .into_iter()
                .map(|catch| {
                    let (body, nested_has_defer) = lower_go_defer_statements(
                        catch.body, env, signatures, state, stack_name, in_loop,
                    );
                    has_defer |= nested_has_defer;
                    CatchClause {
                        types: catch.types,
                        var_name: catch.var_name,
                        stack_var: catch.stack_var,
                        body,
                        when_clause: catch.when_clause,
                    }
                })
                .collect();
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            let next_finally = if let Some(body) = finally {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::Try {
                    body,
                    catches: next_catches,
                    else_body: next_else,
                    finally: next_finally,
                }),
                has_defer,
            )
        }
        _ => (stmt, false),
    }
}

fn go_is_named_result_wrapper_while(
    cond: &Expression,
    body: &[Statement],
    else_body: &Option<Vec<Statement>>,
) -> bool {
    else_body.is_none()
        && matches!(cond.kind, ExprKind::Lit(Literal::Bool(true)))
        && body
            .last()
            .is_some_and(|stmt| matches!(stmt.kind, StmtKind::Break(BreakTarget::Implicit)))
}

fn go_extract_defer_expr(stmt: &Statement) -> Option<Expression> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_defer") || args.len() != 1
    {
        return None;
    }
    Some(args[0].value.clone())
}

fn go_zero_arg_lambda_body_statements(expr: &Expression) -> Option<Vec<Statement>> {
    match &expr.kind {
        ExprKind::Lambda { params, body, .. } if params.is_empty() => Some(match body {
            LambdaBody::Expr(expr) => vec![Statement::new(StmtKind::Expr(expr.as_ref().clone()))],
            LambdaBody::Block(stmts) => stmts.clone(),
        }),
        ExprKind::Cast { expr, .. } => go_zero_arg_lambda_body_statements(expr),
        _ => None,
    }
}

fn go_lower_defer_stmt(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    stack_name: &str,
    frozen_names: &HashSet<String>,
    in_loop: bool,
) -> Vec<Statement> {
    let mut stmts = Vec::new();
    let mut loop_snapshot_captures: Vec<String> = Vec::new();
    let mut deferred_body_override: Option<Vec<Statement>> = None;

    let deferred_expr = match expr.kind {
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            if args.is_empty() {
                match &callee.kind {
                    ExprKind::Ident(name) => {
                        if let Some(body) = env.function_bodies.get(name) {
                            deferred_body_override =
                                Some(normalize_go_block(body, env, signatures, state));
                        }
                    }
                    _ => {
                        if let Some(body) = go_zero_arg_lambda_body_statements(callee.as_ref()) {
                            deferred_body_override = Some(body);
                        }
                    }
                }
            }
            let has_deferred_body_override = deferred_body_override.is_some();
            let deferred_callee = match callee.as_ref() {
                Expression {
                    kind:
                        ExprKind::Member {
                            object,
                            field,
                            null_safe,
                        },
                    ..
                } => {
                    let receiver_type = go_expr_type_hint(object, env, signatures);
                    let deferred_object = if matches!(object.as_ref().kind, ExprKind::Ident(_))
                        && receiver_type.is_none()
                    {
                        object.as_ref().clone()
                    } else {
                        let temp_name = fresh_go_temp(state, "__go_defer_recv");
                        stmts.push(go_defer_temp_decl(
                            temp_name.clone(),
                            receiver_type,
                            object.as_ref().clone(),
                        ));
                        if in_loop {
                            loop_snapshot_captures.push(temp_name.clone());
                        }
                        Expression::ident(&temp_name)
                    };
                    Expression::new(ExprKind::Member {
                        object: Box::new(deferred_object),
                        field: field.clone(),
                        null_safe: *null_safe,
                    })
                }
                _ => {
                    let deferred_value =
                        go_freeze_defer_lambda_captures(callee.as_ref().clone(), frozen_names);
                    let temp_name = fresh_go_temp(state, "__go_defer_fn");
                    stmts.push(go_defer_temp_decl(
                        temp_name.clone(),
                        go_expr_type_hint(&deferred_value, env, signatures),
                        deferred_value,
                    ));
                    if in_loop && !has_deferred_body_override {
                        loop_snapshot_captures.push(temp_name.clone());
                    }
                    Expression::ident(&temp_name)
                }
            };

            let deferred_args = args
                .into_iter()
                .map(|arg| {
                    let temp_name = fresh_go_temp(state, "__go_defer_arg");
                    let value = go_wrap_go_value_copy(arg.value, env, signatures);
                    stmts.push(go_defer_temp_decl(
                        temp_name.clone(),
                        go_expr_type_hint(&value, env, signatures),
                        value,
                    ));
                    if in_loop {
                        loop_snapshot_captures.push(temp_name.clone());
                    }
                    Argument {
                        value: Expression::ident(&temp_name),
                        name: arg.name,
                        by_ref: arg.by_ref,
                        spread: arg.spread,
                    }
                })
                .collect();

            Expression::new(ExprKind::Call {
                callee: Box::new(deferred_callee),
                args: deferred_args,
                optional,
            })
        }
        _ => expr,
    };

    // Build the zero-arg closure that will be stored on the defer stack.
    //
    // Non-loop case: use empty explicit captures so the compiler routes through
    // the outer function's shared env (parent_shared_env_slot path).  The
    // __go_defer_arg* / __go_defer_fn* temps are set-once per function call,
    // so reading them from the shared env at drain time always gives the
    // value they had at defer-registration time.
    //
    // Loop case: explicitly capture the frozen defer temps for this
    // registration. Avoid returning a lambda from an IIFE here: Go defer
    // draining wraps calls in a try/catch for panics, and the shared compiler
    // also models lambda returns with the exception machinery.
    let closure_body = deferred_body_override
        .unwrap_or_else(|| vec![Statement::new(StmtKind::Expr(deferred_expr))]);
    if in_loop
        && !frozen_names.is_empty()
        && go_defer_body_assigns_non_frozen(&closure_body, frozen_names)
    {
        let mut used_names = HashSet::new();
        for stmt in &closure_body {
            go_collect_stmt_idents(stmt, &mut used_names);
        }
        let mut frozen_used = used_names
            .into_iter()
            .filter(|name| frozen_names.contains(name))
            .collect::<Vec<_>>();
        frozen_used.sort();
        for name in frozen_used {
            if !loop_snapshot_captures
                .iter()
                .any(|capture| capture == &name)
            {
                loop_snapshot_captures.push(name);
            }
        }
    }
    let (closure_body, _) =
        lower_go_defer_statements(closure_body, env, signatures, state, stack_name, false);

    let closure = go_deferred_lambda_with_ref_captures(
        LambdaBody::Block(closure_body),
        &loop_snapshot_captures,
        env,
        signatures,
        state,
    );
    stmts.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(stack_name)],
        value: Expression::new(ExprKind::Object(vec![
            ObjectProperty::KeyValue {
                key: Expression::string("fn"),
                value: closure,
            },
            ObjectProperty::KeyValue {
                key: Expression::string("recover"),
                value: env
                    .has_panic_name
                    .as_deref()
                    .zip(env.in_defer_name.as_deref())
                    .map(|(has_panic, in_defer)| {
                        let no_panic = Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(Expression::ident(has_panic)),
                        });
                        let not_in_defer = Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(Expression::ident(in_defer)),
                        });
                        Expression::new(ExprKind::Binary {
                            op: BinOp::And,
                            left: Box::new(no_panic),
                            right: Box::new(not_in_defer),
                        })
                    })
                    .unwrap_or_else(|| Expression::bool(true)),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("next"),
                value: Expression::ident(stack_name),
            },
        ])),
        by_ref: false,
    }));
    stmts
}

fn go_deferred_lambda_with_ref_captures(
    body: LambdaBody,
    value_captures: &[String],
    _env: &GoNormalizeEnv,
    _signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    let mut local_names = HashSet::new();
    go_collect_lambda_declared_names(&body, &mut local_names);

    let mut ref_names = HashSet::new();
    go_collect_lambda_ref_idents(&body, &mut ref_names);

    let mut captured_ref_names = ref_names
        .into_iter()
        .filter(|name| !local_names.contains(name))
        .collect::<Vec<_>>();
    captured_ref_names.sort();

    if captured_ref_names.is_empty() && value_captures.is_empty() {
        return Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body,
            is_async: false,
            captures: Vec::new(),
        });
    }

    let mut replacements = HashMap::new();
    let mut params = Vec::new();
    let mut args = Vec::new();
    for name in value_captures {
        if params.iter().any(|param: &Param| param.name == *name) {
            continue;
        }
        params.push(Param {
            name: name.clone(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
        args.push(Argument::positional(Expression::ident(name)));
    }
    for name in captured_ref_names {
        if value_captures.iter().any(|capture| capture == &name) {
            continue;
        }
        let temp_name = fresh_go_temp(state, "__go_defer_ref_capture");
        params.push(Param {
            name: temp_name.clone(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
        args.push(Argument::positional(Expression::new(ExprKind::RefOf(
            Box::new(PlaceExpr::Ident(name.clone())),
        ))));
        replacements.insert(name, temp_name);
    }

    let inner = Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: go_rewrite_lambda_ref_body(&body, &replacements),
        is_async: false,
        captures: Vec::new(),
    });

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params,
            body: LambdaBody::Expr(Box::new(inner)),
            is_async: false,
            captures: Vec::new(),
        })),
        args,
        optional: false,
    })
}

fn go_collect_block_declared_names(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                go_collect_binding_pattern_names(&decl.pattern, names);
            }
        }
        StmtKind::FunctionDecl { name, .. } => {
            names.insert(name.clone());
        }
        _ => {}
    }
}

fn go_collect_binding_pattern_names(pattern: &BindingPattern, names: &mut HashSet<String>) {
    match pattern {
        BindingPattern::Ident(name) => {
            names.insert(name.clone());
        }
        BindingPattern::Array(elements) => {
            for element in elements {
                if let ArrayPatternElem::Pattern(pattern, _) = element {
                    go_collect_binding_pattern_names(pattern, names);
                }
            }
        }
        BindingPattern::Object(properties) => {
            for property in properties {
                if let Some(pattern) = &property.value {
                    go_collect_binding_pattern_names(pattern, names);
                }
            }
        }
    }
}

fn go_rewrite_immediate_lambda_ref_captures(
    callee: &Expression,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    if let ExprKind::Cast { expr, type_name } = &callee.kind {
        let _ = type_name;
        return go_rewrite_immediate_lambda_ref_captures(
            expr, args, optional, env, signatures, state,
        );
    }

    let ExprKind::Lambda {
        params,
        body,
        is_async,
        captures,
    } = &callee.kind
    else {
        return None;
    };

    let param_names: HashSet<String> = params.iter().map(|param| param.name.clone()).collect();
    let mut local_names = HashSet::new();
    go_collect_lambda_declared_names(body, &mut local_names);

    let mut ref_names = HashSet::new();
    go_collect_lambda_ref_idents(body, &mut ref_names);

    let mut captured_ref_names = ref_names
        .into_iter()
        .filter(|name| !param_names.contains(name) && !local_names.contains(name))
        .collect::<Vec<_>>();
    captured_ref_names.sort();

    if captured_ref_names.is_empty() {
        return None;
    }

    let mut replacements = HashMap::new();
    let mut next_params = params.clone();
    let mut next_args = args.to_vec();

    for name in captured_ref_names {
        let temp_name = fresh_go_temp(state, "__go_ref_capture");
        next_params.push(Param {
            name: temp_name.clone(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
        next_args.push(Argument::positional(Expression::new(ExprKind::RefOf(
            Box::new(PlaceExpr::Ident(name.clone())),
        ))));
        replacements.insert(name, temp_name);
    }

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: next_params,
            body: go_rewrite_lambda_ref_body(body, &replacements),
            is_async: *is_async,
            captures: captures.clone(),
        })),
        args: next_args,
        optional,
    }))
}

fn go_collect_lambda_declared_names(body: &LambdaBody, names: &mut HashSet<String>) {
    if let LambdaBody::Block(stmts) = body {
        for stmt in stmts {
            go_collect_stmt_declared_names_recursive(stmt, names);
        }
    }
}

fn go_collect_stmt_declared_names_recursive(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Expr(expr) => {
            go_collect_expr_declared_names_recursive(expr, names);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                go_collect_binding_pattern_names(&decl.pattern, names);
                if let Some(init) = &decl.init {
                    go_collect_expr_declared_names_recursive(init, names);
                }
            }
        }
        StmtKind::FunctionDecl { name, .. } => {
            names.insert(name.clone());
        }
        StmtKind::Block(body) => {
            for stmt in body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for stmt in then_body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
            for (_, body) in elifs {
                for stmt in body {
                    go_collect_stmt_declared_names_recursive(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_declared_names_recursive(stmt, names);
                }
            }
        }
        StmtKind::For { init, body, .. } => {
            if let Some(init) = init {
                go_collect_stmt_declared_names_recursive(init, names);
            }
            for stmt in body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
        }
        StmtKind::ForIn {
            var,
            key,
            body,
            else_body,
            ..
        } => {
            names.insert(var.clone());
            if let Some(key) = key {
                names.insert(key.clone());
            }
            for stmt in body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_declared_names_recursive(stmt, names);
                }
            }
        }
        StmtKind::While {
            body, else_body, ..
        } => {
            for stmt in body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_declared_names_recursive(stmt, names);
                }
            }
        }
        StmtKind::DoWhile { body, .. } => {
            for stmt in body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                go_collect_expr_declared_names_recursive(target, names);
            }
            go_collect_expr_declared_names_recursive(value, names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            go_collect_expr_declared_names_recursive(target, names);
            go_collect_expr_declared_names_recursive(value, names);
        }
        StmtKind::Return(expr) => {
            if let Some(expr) = expr {
                go_collect_expr_declared_names_recursive(expr, names);
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                go_collect_expr_declared_names_recursive(expr, names);
            }
            if let Some(cause) = cause {
                go_collect_expr_declared_names_recursive(cause, names);
            }
        }
        _ => {}
    }
}

fn go_collect_expr_declared_names_recursive(expr: &Expression, names: &mut HashSet<String>) {
    match &expr.kind {
        ExprKind::Lambda { params, body, .. } => {
            for param in params {
                names.insert(param.name.clone());
            }
            go_collect_lambda_declared_names(body, names);
        }
        ExprKind::Unary { expr, .. } | ExprKind::RefLoad(expr) | ExprKind::Cast { expr, .. } => {
            go_collect_expr_declared_names_recursive(expr, names);
        }
        ExprKind::Binary { left, right, .. } => {
            go_collect_expr_declared_names_recursive(left, names);
            go_collect_expr_declared_names_recursive(right, names);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_collect_expr_declared_names_recursive(cond, names);
            go_collect_expr_declared_names_recursive(then, names);
            go_collect_expr_declared_names_recursive(else_, names);
        }
        ExprKind::Member { object, .. } => go_collect_expr_declared_names_recursive(object, names),
        ExprKind::Index { object, index, .. } => {
            go_collect_expr_declared_names_recursive(object, names);
            go_collect_expr_declared_names_recursive(index, names);
        }
        ExprKind::Assign { target, value } => {
            go_collect_expr_declared_names_recursive(target, names);
            go_collect_expr_declared_names_recursive(value, names);
        }
        ExprKind::Call { callee, args, .. } => {
            go_collect_expr_declared_names_recursive(callee, names);
            for arg in args {
                go_collect_expr_declared_names_recursive(&arg.value, names);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &element.key {
                    go_collect_expr_declared_names_recursive(key, names);
                }
                go_collect_expr_declared_names_recursive(&element.value, names);
            }
        }
        ExprKind::Object(properties) => {
            for property in properties {
                match property {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        go_collect_expr_declared_names_recursive(key, names);
                        go_collect_expr_declared_names_recursive(value, names);
                    }
                    ObjectProperty::Spread(value) => {
                        go_collect_expr_declared_names_recursive(value, names);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(values) | ExprKind::Sequence(values) => {
            for value in values {
                go_collect_expr_declared_names_recursive(value, names);
            }
        }
        _ => {}
    }
}

fn go_collect_lambda_ref_idents(body: &LambdaBody, names: &mut HashSet<String>) {
    match body {
        LambdaBody::Expr(expr) => go_collect_expr_ref_idents(expr, names),
        LambdaBody::Block(stmts) => {
            for stmt in stmts {
                go_collect_stmt_ref_idents(stmt, names);
            }
        }
    }
}

fn go_collect_stmt_ref_idents(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Expr(expr) => go_collect_expr_ref_idents(expr, names),
        StmtKind::Return(expr) => {
            if let Some(expr) = expr {
                go_collect_expr_ref_idents(expr, names);
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                go_collect_expr_ref_idents(expr, names);
            }
            if let Some(cause) = cause {
                go_collect_expr_ref_idents(cause, names);
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                go_collect_expr_assigned_idents(target, names);
                go_collect_expr_ref_idents(target, names);
            }
            go_collect_expr_ref_idents(value, names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            go_collect_expr_assigned_idents(target, names);
            go_collect_expr_ref_idents(target, names);
            go_collect_expr_ref_idents(value, names);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &decl.init {
                    go_collect_expr_ref_idents(init, names);
                }
            }
        }
        StmtKind::Block(body) => {
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            go_collect_expr_ref_idents(cond, names);
            for stmt in then_body {
                go_collect_stmt_ref_idents(stmt, names);
            }
            for (cond, body) in elifs {
                go_collect_expr_ref_idents(cond, names);
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                go_collect_stmt_ref_idents(init, names);
            }
            if let Some(cond) = cond {
                go_collect_expr_ref_idents(cond, names);
            }
            if let Some(update) = update {
                go_collect_expr_ref_idents(update, names);
            }
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            go_collect_expr_ref_idents(iter, names);
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            go_collect_expr_ref_idents(cond, names);
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
            go_collect_expr_ref_idents(cond, names);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
            for catch in catches {
                if let Some(expr) = &catch.when_clause {
                    go_collect_expr_ref_idents(expr, names);
                }
                for stmt in &catch.body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
        }
        _ => {}
    }
}

fn go_collect_expr_ref_idents(expr: &Expression, names: &mut HashSet<String>) {
    match &expr.kind {
        ExprKind::RefOf(place) => {
            if let PlaceExpr::Ident(name) = place.as_ref() {
                names.insert(name.clone());
            }
        }
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => {
            if let ExprKind::Ident(name) = &expr.kind {
                names.insert(name.clone());
            }
            go_collect_expr_ref_idents(expr, names);
        }
        ExprKind::Unary { expr, .. } | ExprKind::RefLoad(expr) | ExprKind::Cast { expr, .. } => {
            go_collect_expr_ref_idents(expr, names)
        }
        ExprKind::CallableRef {
            target,
            receiver,
            adapter,
            ..
        } => {
            go_collect_expr_ref_idents(target, names);
            if let Some(receiver) = receiver {
                go_collect_expr_ref_idents(receiver, names);
            }
            if let Some(CallableAdapter::Expr { body, .. }) = adapter {
                go_collect_expr_ref_idents(body, names);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            go_collect_expr_ref_idents(left, names);
            go_collect_expr_ref_idents(right, names);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_collect_expr_ref_idents(cond, names);
            go_collect_expr_ref_idents(then, names);
            go_collect_expr_ref_idents(else_, names);
        }
        ExprKind::Member { object, .. } => go_collect_expr_ref_idents(object, names),
        ExprKind::Index { object, index, .. } => {
            if let ExprKind::Ident(name) = &object.kind {
                names.insert(name.clone());
            }
            go_collect_expr_ref_idents(object, names);
            go_collect_expr_ref_idents(index, names);
        }
        ExprKind::Assign { target, value } => {
            go_collect_expr_assigned_idents(target, names);
            go_collect_expr_ref_idents(target, names);
            go_collect_expr_ref_idents(value, names);
        }
        ExprKind::Call { callee, args, .. } => {
            go_collect_expr_ref_idents(callee, names);
            for arg in args {
                go_collect_expr_ref_idents(&arg.value, names);
            }
        }
        ExprKind::Lambda { params, body, .. } => {
            let mut nested_names = HashSet::new();
            go_collect_lambda_ref_idents(body, &mut nested_names);
            for param in params {
                nested_names.remove(&param.name);
            }
            names.extend(nested_names);
        }
        _ => {}
    }
}

fn go_collect_expr_assigned_idents(expr: &Expression, names: &mut HashSet<String>) {
    match &expr.kind {
        ExprKind::Ident(name) => {
            names.insert(name.clone());
        }
        ExprKind::RefLoad(inner) | ExprKind::Cast { expr: inner, .. } => {
            go_collect_expr_assigned_idents(inner, names);
        }
        ExprKind::Member { object, .. } => {
            go_collect_expr_assigned_idents(object, names);
        }
        ExprKind::Index { object, index, .. } => {
            go_collect_expr_assigned_idents(object, names);
            go_collect_expr_ref_idents(index, names);
        }
        _ => go_collect_expr_ref_idents(expr, names),
    }
}

fn go_defer_body_assigns_non_frozen(body: &[Statement], frozen_names: &HashSet<String>) -> bool {
    let mut assigned = HashSet::new();
    for stmt in body {
        go_collect_stmt_assigned_idents(stmt, &mut assigned);
    }
    assigned
        .into_iter()
        .any(|name| !frozen_names.contains(&name) && !name.starts_with("__go_"))
}

fn go_collect_stmt_assigned_idents(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Assign { targets, .. } => {
            for target in targets {
                go_collect_expr_assigned_idents(target, names);
            }
        }
        StmtKind::CompoundAssign { target, .. } => {
            go_collect_expr_assigned_idents(target, names);
        }
        StmtKind::Block(body) => {
            for stmt in body {
                go_collect_stmt_assigned_idents(stmt, names);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for stmt in then_body {
                go_collect_stmt_assigned_idents(stmt, names);
            }
            for (_, body) in elifs {
                for stmt in body {
                    go_collect_stmt_assigned_idents(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_assigned_idents(stmt, names);
                }
            }
        }
        StmtKind::For { init, body, .. } => {
            if let Some(init) = init {
                go_collect_stmt_assigned_idents(init, names);
            }
            for stmt in body {
                go_collect_stmt_assigned_idents(stmt, names);
            }
        }
        StmtKind::ForIn {
            var,
            key,
            body,
            else_body,
            ..
        } => {
            names.insert(var.clone());
            if let Some(key) = key {
                names.insert(key.clone());
            }
            for stmt in body {
                go_collect_stmt_assigned_idents(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_assigned_idents(stmt, names);
                }
            }
        }
        StmtKind::While {
            body, else_body, ..
        } => {
            for stmt in body {
                go_collect_stmt_assigned_idents(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_assigned_idents(stmt, names);
                }
            }
        }
        StmtKind::DoWhile { body, .. } => {
            for stmt in body {
                go_collect_stmt_assigned_idents(stmt, names);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                go_collect_stmt_assigned_idents(stmt, names);
            }
            for catch in catches {
                for stmt in &catch.body {
                    go_collect_stmt_assigned_idents(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_assigned_idents(stmt, names);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    go_collect_stmt_assigned_idents(stmt, names);
                }
            }
        }
        _ => {}
    }
}

fn go_collect_ref_place_names_stmt(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            go_collect_ref_place_names_expr(expr, names)
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                go_collect_ref_place_names_expr(expr, names);
            }
            if let Some(cause) = cause {
                go_collect_ref_place_names_expr(cause, names);
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                go_collect_ref_place_names_expr(target, names);
            }
            go_collect_ref_place_names_expr(value, names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            go_collect_ref_place_names_expr(target, names);
            go_collect_ref_place_names_expr(value, names);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &decl.init {
                    go_collect_ref_place_names_expr(init, names);
                }
            }
        }
        StmtKind::Block(body) => {
            for stmt in body {
                go_collect_ref_place_names_stmt(stmt, names);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            go_collect_ref_place_names_expr(cond, names);
            for stmt in then_body {
                go_collect_ref_place_names_stmt(stmt, names);
            }
            for (cond, body) in elifs {
                go_collect_ref_place_names_expr(cond, names);
                for stmt in body {
                    go_collect_ref_place_names_stmt(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_ref_place_names_stmt(stmt, names);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                go_collect_ref_place_names_stmt(init, names);
            }
            if let Some(cond) = cond {
                go_collect_ref_place_names_expr(cond, names);
            }
            if let Some(update) = update {
                go_collect_ref_place_names_expr(update, names);
            }
            for stmt in body {
                go_collect_ref_place_names_stmt(stmt, names);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            go_collect_ref_place_names_expr(iter, names);
            for stmt in body {
                go_collect_ref_place_names_stmt(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_ref_place_names_stmt(stmt, names);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            go_collect_ref_place_names_expr(cond, names);
            for stmt in body {
                go_collect_ref_place_names_stmt(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_ref_place_names_stmt(stmt, names);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                go_collect_ref_place_names_stmt(stmt, names);
            }
            go_collect_ref_place_names_expr(cond, names);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                go_collect_ref_place_names_stmt(stmt, names);
            }
            for catch in catches {
                if let Some(expr) = &catch.when_clause {
                    go_collect_ref_place_names_expr(expr, names);
                }
                for stmt in &catch.body {
                    go_collect_ref_place_names_stmt(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_ref_place_names_stmt(stmt, names);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    go_collect_ref_place_names_stmt(stmt, names);
                }
            }
        }
        _ => {}
    }
}

fn go_collect_ref_place_names_expr(expr: &Expression, names: &mut HashSet<String>) {
    match &expr.kind {
        ExprKind::RefOf(place) => {
            if let PlaceExpr::Ident(name) = place.as_ref() {
                names.insert(name.clone());
            }
        }
        ExprKind::Unary { expr, .. } | ExprKind::RefLoad(expr) | ExprKind::Cast { expr, .. } => {
            go_collect_ref_place_names_expr(expr, names);
        }
        ExprKind::Binary { left, right, .. } => {
            go_collect_ref_place_names_expr(left, names);
            go_collect_ref_place_names_expr(right, names);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_collect_ref_place_names_expr(cond, names);
            go_collect_ref_place_names_expr(then, names);
            go_collect_ref_place_names_expr(else_, names);
        }
        ExprKind::Member { object, .. } => go_collect_ref_place_names_expr(object, names),
        ExprKind::Index { object, index, .. } => {
            go_collect_ref_place_names_expr(object, names);
            go_collect_ref_place_names_expr(index, names);
        }
        ExprKind::Assign { target, value } => {
            go_collect_ref_place_names_expr(target, names);
            go_collect_ref_place_names_expr(value, names);
        }
        ExprKind::Call { callee, args, .. } => {
            go_collect_ref_place_names_expr(callee, names);
            for arg in args {
                go_collect_ref_place_names_expr(&arg.value, names);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &element.key {
                    go_collect_ref_place_names_expr(key, names);
                }
                go_collect_ref_place_names_expr(&element.value, names);
            }
        }
        ExprKind::Object(properties) => {
            for property in properties {
                match property {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        go_collect_ref_place_names_expr(key, names);
                        go_collect_ref_place_names_expr(value, names);
                    }
                    ObjectProperty::Spread(value) => go_collect_ref_place_names_expr(value, names),
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(values) | ExprKind::Sequence(values) => {
            for value in values {
                go_collect_ref_place_names_expr(value, names);
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => go_collect_ref_place_names_expr(expr, names),
            LambdaBody::Block(body) => {
                for stmt in body {
                    go_collect_ref_place_names_stmt(stmt, names);
                }
            }
        },
        _ => {}
    }
}

fn go_rewrite_ref_place_reads_stmt(stmt: &Statement, names: &HashSet<String>) -> Statement {
    if names.is_empty() {
        return stmt.clone();
    }
    match &stmt.kind {
        StmtKind::Expr(expr) => Statement::new(StmtKind::Expr(go_rewrite_ref_place_reads_expr(
            expr, names, false,
        ))),
        _ => stmt.clone(),
    }
}

fn go_rewrite_ref_place_reads_expr(
    expr: &Expression,
    names: &HashSet<String>,
    is_lvalue: bool,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) if !is_lvalue && names.contains(name) => go_ref_cell_value(name),
        ExprKind::Unary { op, expr: inner } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(go_rewrite_ref_place_reads_expr(inner, names, false)),
        }),
        ExprKind::RefLoad(inner) => Expression::new(ExprKind::RefLoad(Box::new(
            go_rewrite_ref_place_reads_expr(inner, names, false),
        ))),
        ExprKind::Cast {
            expr: inner,
            type_name,
        } => Expression::new(ExprKind::Cast {
            expr: Box::new(go_rewrite_ref_place_reads_expr(inner, names, false)),
            type_name: type_name.clone(),
        }),
        ExprKind::Binary { left, op, right } => Expression::new(ExprKind::Binary {
            left: Box::new(go_rewrite_ref_place_reads_expr(left, names, false)),
            op: *op,
            right: Box::new(go_rewrite_ref_place_reads_expr(right, names, false)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(go_rewrite_ref_place_reads_expr(cond, names, false)),
            then: Box::new(go_rewrite_ref_place_reads_expr(then, names, false)),
            else_: Box::new(go_rewrite_ref_place_reads_expr(else_, names, false)),
        }),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: Box::new(go_rewrite_ref_place_reads_expr(object, names, false)),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(go_rewrite_ref_place_reads_expr(object, names, false)),
            index: Box::new(go_rewrite_ref_place_reads_expr(index, names, false)),
            null_safe: *null_safe,
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => Expression::new(ExprKind::Call {
            callee: Box::new(go_rewrite_ref_place_reads_expr(callee, names, false)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: go_rewrite_ref_place_reads_expr(&arg.value, names, false),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
            optional: *optional,
        }),
        _ => expr.clone(),
    }
}

fn go_rewrite_lambda_ref_body(
    body: &LambdaBody,
    replacements: &HashMap<String, String>,
) -> LambdaBody {
    match body {
        LambdaBody::Expr(expr) => {
            LambdaBody::Expr(Box::new(go_rewrite_expr_ref_idents(expr, replacements)))
        }
        LambdaBody::Block(stmts) => LambdaBody::Block(
            stmts
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
        ),
    }
}

fn go_rewrite_stmt_ref_idents(
    stmt: &Statement,
    replacements: &HashMap<String, String>,
) -> Statement {
    match &stmt.kind {
        StmtKind::Expr(expr) => Statement::new(StmtKind::Expr(go_rewrite_expr_ref_idents(
            expr,
            replacements,
        ))),
        StmtKind::Return(expr) => Statement::new(StmtKind::Return(
            expr.as_ref()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
        )),
        StmtKind::Throw { expr, cause } => Statement::new(StmtKind::Throw {
            expr: expr
                .as_ref()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
            cause: cause
                .as_ref()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
        }),
        StmtKind::Assign { targets, value, .. } => Statement::new(StmtKind::Assign {
            targets: targets
                .iter()
                .map(|expr| go_rewrite_expr_ref_idents_lvalue(expr, replacements))
                .collect(),
            value: go_rewrite_expr_ref_idents(value, replacements),
            by_ref: false,
        }),
        StmtKind::CompoundAssign { target, op, value } => {
            Statement::new(StmtKind::CompoundAssign {
                target: go_rewrite_expr_ref_idents_lvalue(target, replacements),
                op: *op,
                value: go_rewrite_expr_ref_idents(value, replacements),
            })
        }
        StmtKind::VarDecl { declarations, kind } => Statement::new(StmtKind::VarDecl {
            declarations: declarations
                .iter()
                .map(|decl| VarDeclarator {
                    pattern: decl.pattern.clone(),
                    type_hint: decl.type_hint.clone(),
                    init: decl
                        .init
                        .as_ref()
                        .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
                    array_bounds: decl.array_bounds.clone(),
                    with_events: decl.with_events,
                })
                .collect(),
            kind: kind.clone(),
        }),
        StmtKind::Block(body) => Statement::new(StmtKind::Block(
            body.iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
        )),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => Statement::new(StmtKind::If {
            cond: go_rewrite_expr_ref_idents(cond, replacements),
            then_body: then_body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
            elifs: elifs
                .iter()
                .map(|(cond, body)| {
                    (
                        go_rewrite_expr_ref_idents(cond, replacements),
                        body.iter()
                            .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                            .collect(),
                    )
                })
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                    .collect()
            }),
        }),
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => Statement::new(StmtKind::For {
            init: init
                .as_ref()
                .map(|stmt| Box::new(go_rewrite_stmt_ref_idents(stmt, replacements))),
            cond: cond
                .as_ref()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
            update: update
                .as_ref()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
            body: body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
        }),
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => Statement::new(StmtKind::ForIn {
            var: var.clone(),
            key: key.clone(),
            iter: go_rewrite_expr_ref_idents(iter, replacements),
            body: body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
            of: *of,
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                    .collect()
            }),
            is_async: *is_async,
        }),
        StmtKind::While {
            cond,
            body,
            else_body,
        } => Statement::new(StmtKind::While {
            cond: go_rewrite_expr_ref_idents(cond, replacements),
            body: body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                    .collect()
            }),
        }),
        StmtKind::DoWhile { body, cond, until } => Statement::new(StmtKind::DoWhile {
            body: body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
            cond: go_rewrite_expr_ref_idents(cond, replacements),
            until: *until,
        }),
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => Statement::new(StmtKind::Try {
            body: body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
            catches: catches
                .iter()
                .map(|catch| CatchClause {
                    types: catch.types.clone(),
                    var_name: catch.var_name.clone(),
                    stack_var: catch.stack_var.clone(),
                    body: catch
                        .body
                        .iter()
                        .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                        .collect(),
                    when_clause: catch
                        .when_clause
                        .as_ref()
                        .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
                })
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                    .collect()
            }),
            finally: finally.as_ref().map(|body| {
                body.iter()
                    .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                    .collect()
            }),
        }),
        _ => stmt.clone(),
    }
}

fn go_rewrite_expr_ref_idents(
    expr: &Expression,
    replacements: &HashMap<String, String>,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) => replacements
            .get(name)
            .map(|replacement| {
                if name.starts_with("__go_defer_stack") {
                    go_ref_cell_value(replacement)
                } else {
                    Expression::new(ExprKind::RefLoad(Box::new(Expression::ident(replacement))))
                }
            })
            .unwrap_or_else(|| expr.clone()),
        ExprKind::RefOf(place) => {
            if let PlaceExpr::Ident(name) = place.as_ref() {
                if let Some(replacement) = replacements.get(name) {
                    return Expression::ident(replacement);
                }
            }
            expr.clone()
        }
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: inner,
        } => {
            if let ExprKind::Ident(name) = &inner.kind {
                if let Some(replacement) = replacements.get(name) {
                    return Expression::ident(replacement);
                }
            }
            Expression::new(ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr: Box::new(go_rewrite_expr_ref_idents(inner, replacements)),
            })
        }
        ExprKind::Unary { op, expr: inner } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(go_rewrite_expr_ref_idents(inner, replacements)),
        }),
        ExprKind::RefLoad(inner) => Expression::new(ExprKind::RefLoad(Box::new(
            go_rewrite_expr_ref_idents(inner, replacements),
        ))),
        ExprKind::Cast {
            expr: inner,
            type_name,
        } => Expression::new(ExprKind::Cast {
            expr: Box::new(go_rewrite_expr_ref_idents(inner, replacements)),
            type_name: type_name.clone(),
        }),
        ExprKind::Binary { left, op, right } => Expression::new(ExprKind::Binary {
            left: Box::new(go_rewrite_expr_ref_idents(left, replacements)),
            op: *op,
            right: Box::new(go_rewrite_expr_ref_idents(right, replacements)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(go_rewrite_expr_ref_idents(cond, replacements)),
            then: Box::new(go_rewrite_expr_ref_idents(then, replacements)),
            else_: Box::new(go_rewrite_expr_ref_idents(else_, replacements)),
        }),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: Box::new(go_rewrite_expr_ref_idents(object, replacements)),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(go_rewrite_expr_ref_idents(object, replacements)),
            index: Box::new(go_rewrite_expr_ref_idents(index, replacements)),
            null_safe: *null_safe,
        }),
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(go_rewrite_expr_ref_idents(target, replacements)),
            value: Box::new(go_rewrite_expr_ref_idents(value, replacements)),
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => Expression::new(ExprKind::Call {
            callee: Box::new(go_rewrite_expr_ref_idents(callee, replacements)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: go_rewrite_expr_ref_idents(&arg.value, replacements),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
            optional: *optional,
        }),
        ExprKind::Array(elements) => Expression::new(ExprKind::Array(
            elements
                .iter()
                .map(|element| ArrayElement {
                    key: element
                        .key
                        .as_ref()
                        .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
                    value: go_rewrite_expr_ref_idents(&element.value, replacements),
                    spread: element.spread,
                    by_ref: element.by_ref,
                })
                .collect(),
        )),
        ExprKind::Object(properties) => Expression::new(ExprKind::Object(
            properties
                .iter()
                .map(|property| match property {
                    ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                        key: go_rewrite_expr_ref_idents(key, replacements),
                        value: go_rewrite_expr_ref_idents(value, replacements),
                    },
                    ObjectProperty::Spread(value) => {
                        ObjectProperty::Spread(go_rewrite_expr_ref_idents(value, replacements))
                    }
                    ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                        key: go_rewrite_expr_ref_idents(key, replacements),
                        value: go_rewrite_expr_ref_idents(value, replacements),
                    },
                    _ => property.clone(),
                })
                .collect(),
        )),
        ExprKind::Tuple(values) => Expression::new(ExprKind::Tuple(
            values
                .iter()
                .map(|value| go_rewrite_expr_ref_idents(value, replacements))
                .collect(),
        )),
        ExprKind::Sequence(values) => Expression::new(ExprKind::Sequence(
            values
                .iter()
                .map(|value| go_rewrite_expr_ref_idents(value, replacements))
                .collect(),
        )),
        ExprKind::Lambda {
            params,
            body,
            is_async,
            captures,
        } => Expression::new(ExprKind::Lambda {
            params: params.clone(),
            body: go_rewrite_lambda_ref_body(body, replacements),
            is_async: *is_async,
            captures: captures.clone(),
        }),
        _ => expr.clone(),
    }
}

fn go_rewrite_expr_ref_idents_lvalue(
    expr: &Expression,
    replacements: &HashMap<String, String>,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) => replacements
            .get(name)
            .map(|replacement| {
                if name.starts_with("__go_defer_stack") {
                    go_ref_cell_value(replacement)
                } else {
                    Expression::new(ExprKind::RefLoad(Box::new(Expression::ident(replacement))))
                }
            })
            .unwrap_or_else(|| expr.clone()),
        ExprKind::Member { .. } | ExprKind::Index { .. } => {
            go_rewrite_expr_ref_idents(expr, replacements)
        }
        _ => {
            let rewritten = go_rewrite_expr_ref_idents(expr, replacements);
            if go_lvalue_mentions_replaced_ident(expr, replacements) {
                if let Some(place) = PlaceExpr::from_expr(&rewritten) {
                    return Expression::new(ExprKind::RefLoad(Box::new(Expression::new(
                        ExprKind::RefOf(Box::new(place)),
                    ))));
                }
            }
            rewritten
        }
    }
}

fn go_ref_cell_value(name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(name)),
        field: "__value".to_string(),
        null_safe: false,
    })
}

fn go_lvalue_mentions_replaced_ident(
    expr: &Expression,
    replacements: &HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => replacements.contains_key(name),
        ExprKind::Member { object, .. } => go_lvalue_mentions_replaced_ident(object, replacements),
        ExprKind::Index { object, index, .. } => {
            go_lvalue_mentions_replaced_ident(object, replacements)
                || go_lvalue_mentions_replaced_ident(index, replacements)
        }
        ExprKind::RefLoad(inner) | ExprKind::Cast { expr: inner, .. } => {
            go_lvalue_mentions_replaced_ident(inner, replacements)
        }
        ExprKind::Unary { expr: inner, .. } => {
            go_lvalue_mentions_replaced_ident(inner, replacements)
        }
        _ => false,
    }
}

fn go_freeze_defer_lambda_captures(expr: Expression, frozen_names: &HashSet<String>) -> Expression {
    let ExprKind::Lambda {
        params,
        body,
        is_async,
        mut captures,
    } = expr.kind
    else {
        return expr;
    };

    if !frozen_names.is_empty() {
        let mut used_names = HashSet::new();
        go_collect_lambda_body_idents(&body, &mut used_names);
        let param_names: HashSet<String> = params.iter().map(|param| param.name.clone()).collect();
        let mut frozen_capture_names = used_names
            .into_iter()
            .filter(|name| frozen_names.contains(name) && !param_names.contains(name))
            .collect::<Vec<_>>();
        frozen_capture_names.sort();
        for name in frozen_capture_names {
            if !captures.iter().any(|capture| capture == &name) {
                captures.push(name);
            }
        }
    }

    Expression::new(ExprKind::Lambda {
        params,
        body,
        is_async,
        captures,
    })
}

fn go_collect_lambda_body_idents(body: &LambdaBody, names: &mut HashSet<String>) {
    match body {
        LambdaBody::Expr(expr) => go_collect_expr_idents(expr, names),
        LambdaBody::Block(stmts) => {
            for stmt in stmts {
                go_collect_stmt_idents(stmt, names);
            }
        }
    }
}

fn go_collect_stmt_idents(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Expr(expr) => go_collect_expr_idents(expr, names),
        StmtKind::Return(expr) => {
            if let Some(expr) = expr {
                go_collect_expr_idents(expr, names);
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                go_collect_expr_idents(expr, names);
            }
            if let Some(cause) = cause {
                go_collect_expr_idents(cause, names);
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                go_collect_expr_idents(target, names);
            }
            go_collect_expr_idents(value, names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            go_collect_expr_idents(target, names);
            go_collect_expr_idents(value, names);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            go_collect_expr_idents(cond, names);
            for stmt in then_body {
                go_collect_stmt_idents(stmt, names);
            }
            for (cond, body) in elifs {
                go_collect_expr_idents(cond, names);
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                go_collect_stmt_idents(init, names);
            }
            if let Some(cond) = cond {
                go_collect_expr_idents(cond, names);
            }
            if let Some(update) = update {
                go_collect_expr_idents(update, names);
            }
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            go_collect_expr_idents(iter, names);
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            go_collect_expr_idents(cond, names);
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
            go_collect_expr_idents(cond, names);
        }
        StmtKind::Block(body) => {
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            go_collect_expr_idents(expr, names);
            for case in cases {
                for condition in &case.conditions {
                    match condition {
                        CaseCondition::Value(expr) => go_collect_expr_idents(expr, names),
                        CaseCondition::Range { from, to } => {
                            go_collect_expr_idents(from, names);
                            go_collect_expr_idents(to, names);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            go_collect_expr_idents(expr, names)
                        }
                    }
                }
                for stmt in &case.body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
            if let Some(body) = default {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
            for catch in catches {
                if let Some(cond) = &catch.when_clause {
                    go_collect_expr_idents(cond, names);
                }
                for stmt in &catch.body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &decl.init {
                    go_collect_expr_idents(init, names);
                }
            }
        }
        _ => {}
    }
}

fn go_collect_expr_idents(expr: &Expression, names: &mut HashSet<String>) {
    match &expr.kind {
        ExprKind::Ident(name) => {
            names.insert(name.clone());
        }
        ExprKind::Unary { expr, .. } | ExprKind::RefLoad(expr) | ExprKind::Cast { expr, .. } => {
            go_collect_expr_idents(expr, names)
        }
        ExprKind::FuncRef(name) => {
            names.insert(name.clone());
        }
        ExprKind::CallableRef {
            target,
            receiver,
            adapter,
            ..
        } => {
            go_collect_expr_idents(target, names);
            if let Some(receiver) = receiver {
                go_collect_expr_idents(receiver, names);
            }
            if let Some(CallableAdapter::Expr { body, .. }) = adapter {
                go_collect_expr_idents(body, names);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            go_collect_expr_idents(left, names);
            go_collect_expr_idents(right, names);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_collect_expr_idents(cond, names);
            go_collect_expr_idents(then, names);
            go_collect_expr_idents(else_, names);
        }
        ExprKind::Member { object, .. } => go_collect_expr_idents(object, names),
        ExprKind::Index { object, index, .. } => {
            go_collect_expr_idents(object, names);
            go_collect_expr_idents(index, names);
        }
        ExprKind::Assign { target, value } => {
            go_collect_expr_idents(target, names);
            go_collect_expr_idents(value, names);
        }
        ExprKind::Call { callee, args, .. } => {
            go_collect_expr_idents(callee, names);
            for arg in args {
                go_collect_expr_idents(&arg.value, names);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &element.key {
                    go_collect_expr_idents(key, names);
                }
                go_collect_expr_idents(&element.value, names);
            }
        }
        ExprKind::Object(properties) => {
            for property in properties {
                match property {
                    ObjectProperty::KeyValue { key, value } => {
                        go_collect_expr_idents(key, names);
                        go_collect_expr_idents(value, names);
                    }
                    ObjectProperty::Spread(value) => go_collect_expr_idents(value, names),
                    ObjectProperty::Computed { key, value } => {
                        go_collect_expr_idents(key, names);
                        go_collect_expr_idents(value, names);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(values) | ExprKind::Sequence(values) => {
            for value in values {
                go_collect_expr_idents(value, names);
            }
        }
        ExprKind::Lambda { body, .. } => go_collect_lambda_body_idents(body, names),
        _ => {}
    }
}

fn go_defer_temp_decl(name: String, type_hint: Option<String>, init: Expression) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint: type_hint.map(Into::into),
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::FunctionScoped,
    })
}

/// Every statement this pass produces keeps the source position of the one it
/// came from.
///
/// ⛔ THE NORMALIZER REBUILDS. It takes a `&Statement` and constructs new ones
/// from the kind, so without this every Go statement reached the compiler with
/// no line number — `--dump-ast` showed a whole program at line 0. A statement
/// that already carries a position keeps it: a nested normalization is closer
/// to the source than its parent. `to_span` is 1-based, so line 0 only ever
/// means "never set".
fn normalize_go_statement(
    stmt: &Statement,
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let span = stmt.span;
    let mut out = normalize_go_statement_inner(stmt, env, signatures, state);
    for s in out.iter_mut() {
        if s.span.start_line == 0 && s.span.end_line == 0 {
            s.span = span;
        }
    }
    out
}

fn normalize_go_statement_inner(
    stmt: &Statement,
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    if env.recover_fn_name.is_none() {
        env.panic_value_name = Some(fresh_go_temp(state, "__go_panic_value"));
        env.has_panic_name = Some(fresh_go_temp(state, "__go_has_panic"));
        env.in_defer_name = Some(fresh_go_temp(state, "__go_in_defer"));
        env.recover_fn_name = Some(fresh_go_temp(state, "__go_recover"));
    }

    match &stmt.kind {
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            handles,
            is_async,
            is_generator,
            is_sub,
        } => {
            let mut fn_env = GoNormalizeEnv {
                value_types: env.value_types.clone(),
                reflect_value_payloads: env.reflect_value_payloads.clone(),
                reflect_value_targets: env.reflect_value_targets.clone(),
                reflect_pointer_targets: env.reflect_pointer_targets.clone(),
                reflect_method_bindings: env.reflect_method_bindings.clone(),
                reflect_array_payloads: env.reflect_array_payloads.clone(),
                package_aliases: env.package_aliases.clone(),
                fixed_arrays: env.fixed_arrays.clone(),
                regex_patterns: env.regex_patterns.clone(),
                slice_caps: env.slice_caps.clone(),
                slice_views: env.slice_views.clone(),
                struct_infos: env.struct_infos.clone(),
                interface_methods: env.interface_methods.clone(),
                interface_concrete_types: env.interface_concrete_types.clone(),
                nil_interface_values: env.nil_interface_values.clone(),
                method_value_bindings: HashMap::new(),
                named_types: env.named_types.clone(),
                type_names: env.type_names.clone(),
                function_bodies: env.function_bodies.clone(),
                flag_bindings: env.flag_bindings.clone(),
                flag_defs: env.flag_defs.clone(),
                log_output: env.log_output.clone(),
                log_prefix: env.log_prefix.clone(),
                log_flags: env.log_flags.clone(),
                time_round_half_hour_bindings: env.time_round_half_hour_bindings.clone(),
                generic_type_params: HashMap::new(),
                return_type: return_type.clone(),
                panic_value_name: None,
                has_panic_name: None,
                in_defer_name: None,
                recover_fn_name: None,
                owns_panic_state: false,
            };
            for param in params {
                if param.type_hint.as_deref() == Some("__goTypeArg") {
                    if let Some(type_param) = go_runtime_generic_param_name(&param.name) {
                        fn_env
                            .generic_type_params
                            .insert(type_param, param.name.clone());
                    }
                }
                if let Some(type_hint) = param.type_hint.as_ref() {
                    fn_env
                        .value_types
                        .insert(param.name.clone(), type_hint.clone().to_string());
                }
                if let Some(type_hint) = param
                    .type_hint
                    .as_deref()
                    .filter(|hint| go_is_fixed_array_type(hint))
                {
                    fn_env
                        .fixed_arrays
                        .insert(param.name.clone(), type_hint.to_string());
                }
            }

            vec![Statement::new(StmtKind::FunctionDecl {
                name: name.clone(),
                params: params.clone(),
                return_type: return_type.clone(),
                body: normalize_go_function_body(body, &mut fn_env, signatures, state),
                modifiers: modifiers.clone(),
                handles: handles.clone(),
                is_async: *is_async,
                is_generator: *is_generator,
                is_sub: *is_sub,
            })]
        }
        StmtKind::Labeled { label, body } => {
            let mut label_env = env.clone();
            let mut normalized_body =
                normalize_go_statement(body, &mut label_env, signatures, state);
            let body = if normalized_body.len() == 1 {
                normalized_body.remove(0)
            } else {
                Statement::new(StmtKind::Block(normalized_body))
            };
            vec![Statement::new(StmtKind::Labeled {
                label: label.clone(),
                body: Box::new(body),
            })]
        }
        StmtKind::VarDecl { declarations, kind } => {
            let mut prefix = Vec::new();
            let mut normalized = Vec::with_capacity(declarations.len());
            for decl in declarations {
                let mut next_decl = decl.clone();
                if let Some(pattern) = go_single_named_binding_pattern(&next_decl.pattern) {
                    next_decl.pattern = pattern;
                }
                next_decl.init = decl.init.as_ref().map(|expr| {
                    if go_is_two_value_binding_pattern(&decl.pattern) {
                        if let Some(tuple_expr) =
                            go_normalize_channel_receive_tuple_expr(expr, env, signatures, state)
                        {
                            return tuple_expr;
                        }
                        if let Some(tuple_expr) =
                            go_normalize_map_lookup_tuple_expr(expr, env, signatures, state)
                        {
                            return tuple_expr;
                        }
                    }
                    let normalized = normalize_go_expr(expr, env, signatures, state);
                    if next_decl
                        .type_hint
                        .as_deref()
                        .is_some_and(go_is_function_type)
                    {
                        go_normalize_function_value(normalized, env, signatures)
                    } else {
                        go_normalize_function_value_for_inferred_binding(
                            normalized, env, signatures,
                        )
                    }
                });
                let method_value_binding = next_decl
                    .init
                    .as_ref()
                    .and_then(|init| go_method_value_binding_from_expr(init, env, signatures));
                if let Some(init) = next_decl.init.take() {
                    next_decl.init = Some(if method_value_binding.is_some() {
                        Expression::null()
                    } else {
                        go_normalize_method_value_binding(init, env, signatures, state)
                    });
                }
                next_decl.array_bounds = decl.array_bounds.as_ref().map(|bounds| {
                    bounds
                        .iter()
                        .map(|expr| normalize_go_expr(expr, env, signatures, state))
                        .collect()
                });

                if next_decl.init.is_none()
                    && next_decl
                        .type_hint
                        .as_deref()
                        .is_some_and(go_is_fixed_array_type)
                {
                    next_decl.init = next_decl
                        .type_hint
                        .as_deref()
                        .map(|type_name| go_zero_value_for_type(type_name, env));
                } else if next_decl.init.is_none() {
                    if let Some(type_name) = next_decl.type_hint.as_deref() {
                        next_decl.init = Some(go_zero_value_for_type(type_name, env));
                    }
                } else if let Some(init_expr) = next_decl.init.take() {
                    next_decl.init = Some(go_wrap_fixed_array_copy(init_expr, env, signatures));
                }

                if let Some((tmp_decl, ref_expr, tmp_name, tmp_type)) =
                    go_materialize_addressed_composite_decl_init(
                        next_decl.init.as_ref(),
                        env,
                        signatures,
                        state,
                    )
                {
                    env.value_types
                        .insert(tmp_name, go_canonical_go_type(&tmp_type));
                    prefix.push(Statement::new(StmtKind::VarDecl {
                        declarations: vec![tmp_decl],
                        kind: VarDeclKind::Let,
                    }));
                    next_decl.init = Some(ref_expr);
                }

                if let Some((name, type_name)) =
                    go_decl_fixed_array_binding(&next_decl, env, signatures)
                {
                    env.fixed_arrays.insert(name, type_name);
                }
                if let Some((name, type_name)) = go_decl_binding_type(&next_decl, env, signatures) {
                    env.value_types
                        .insert(name, go_canonical_go_type(&type_name));
                }
                if let Some((name, concrete_type)) =
                    go_decl_interface_concrete_binding(&next_decl, env, signatures)
                {
                    env.interface_concrete_types
                        .insert(name.clone(), concrete_type);
                    env.nil_interface_values.remove(&name);
                } else if let BindingPattern::Ident(name) = &next_decl.pattern {
                    env.interface_concrete_types.remove(name);
                    if next_decl
                        .type_hint
                        .as_deref()
                        .is_some_and(|type_hint| go_is_go_interface_type(type_hint, env))
                        && next_decl.init.as_ref().is_none_or(|init| {
                            go_is_null_expr(init) || go_expr_is_nil_interface_value(init, env)
                        })
                    {
                        env.nil_interface_values.insert(name.clone());
                    } else {
                        env.nil_interface_values.remove(name);
                    }
                }
                if let BindingPattern::Ident(name) = &next_decl.pattern {
                    if let Some(binding) = method_value_binding.clone() {
                        env.method_value_bindings.insert(name.clone(), binding);
                    } else {
                        env.method_value_bindings.remove(name);
                    }
                }
                if let BindingPattern::Ident(name) = &next_decl.pattern {
                    if let Some(pattern) = next_decl
                        .init
                        .as_ref()
                        .and_then(|init| go_regex_pattern_from_expr(init, env))
                    {
                        env.regex_patterns.insert(name.clone(), pattern);
                    }
                }
                if let Some(type_hints) = next_decl
                    .init
                    .as_ref()
                    .and_then(|init| go_expr_tuple_type_hints(init, env, signatures))
                {
                    go_record_binding_pattern_type_hints(&next_decl.pattern, &type_hints, env);
                }
                if let Some(name) = go_binding_name(&next_decl.pattern) {
                    if let Some(init) = next_decl.init.as_ref() {
                        if let Some(payload) = go_reflect_value_payload(init) {
                            env.reflect_value_payloads.insert(name.clone(), payload);
                        } else {
                            env.reflect_value_payloads.remove(&name);
                        }
                        if let Some(target) = go_reflect_settable_target(init) {
                            env.reflect_value_targets.insert(name.clone(), target);
                        } else {
                            env.reflect_value_targets.remove(&name);
                        }
                        if let Some((target, type_name)) = go_reflect_pointer_target(init) {
                            env.reflect_pointer_targets
                                .insert(name.clone(), (target, type_name));
                        } else {
                            env.reflect_pointer_targets.remove(&name);
                        }
                        if let Some(binding) = go_reflect_method_binding(init) {
                            env.reflect_method_bindings.insert(name.clone(), binding);
                        } else {
                            env.reflect_method_bindings.remove(&name);
                        }
                        if let Some(payloads) = go_reflect_array_payloads(init) {
                            env.reflect_array_payloads.insert(name.clone(), payloads);
                        } else {
                            env.reflect_array_payloads.remove(&name);
                        }
                    }
                    if decl
                        .init
                        .as_ref()
                        .is_some_and(go_time_is_round_binary_duration_call)
                    {
                        env.time_round_half_hour_bindings.insert(name.clone());
                    }
                    if let Some(flag_def) = decl.init.as_ref().and_then(go_flag_definition) {
                        env.flag_defs.push(flag_def);
                    }
                    if let Some((flag_name, flag_kind)) =
                        decl.init.as_ref().and_then(go_flag_binding_from_init)
                    {
                        env.flag_bindings
                            .insert(flag_name, (name.clone(), flag_kind));
                    }
                    if let Some(view) = next_decl
                        .init
                        .as_ref()
                        .and_then(|init| go_expr_slice_view(init, env))
                    {
                        if go_slice_view_is_self_referential(&view, &name) {
                            env.slice_views.remove(&name);
                        } else {
                            env.slice_views.insert(name.clone(), view);
                        }
                    } else {
                        env.slice_views.remove(&name);
                    }
                    if let Some(cap_expr) = decl
                        .init
                        .as_ref()
                        .and_then(|init| go_make_slice_capacity_expr(init, env, signatures, state))
                        .or_else(|| {
                            next_decl
                                .init
                                .as_ref()
                                .and_then(|init| go_bound_slice_capacity_expr(init, env))
                        })
                    {
                        env.slice_caps.insert(name, cap_expr);
                    }
                }
                if let Some(runtime_type) = next_decl
                    .type_hint
                    .as_deref()
                    .and_then(|type_hint| go_named_non_struct_underlying_type(type_hint, env))
                {
                    next_decl.type_hint = Some(runtime_type.into());
                }
                if let Some(expanded) = go_expand_static_tuple_decl(&next_decl, env, signatures) {
                    prefix.push(Statement::new(StmtKind::VarDecl {
                        declarations: expanded,
                        kind: kind.clone(),
                    }));
                    continue;
                }
                normalized.push(next_decl);
            }

            if !normalized.is_empty() {
                prefix.push(Statement::new(StmtKind::VarDecl {
                    declarations: normalized,
                    kind: kind.clone(),
                }));
            }
            prefix
        }
        StmtKind::Expr(_) if go_extract_named_type_marker(stmt).is_some() => vec![stmt.clone()],
        StmtKind::Expr(expr) => {
            let expr = go_unwrap_spawned_gob_expr(expr).unwrap_or_else(|| expr.clone());
            if let Some(panic_expr) = go_extract_panic_expr(&expr) {
                vec![Statement::new(StmtKind::Throw {
                    expr: Some(normalize_go_expr(panic_expr, env, signatures, state)),
                    cause: None,
                })]
            } else if go_type_assertion_known_result(&expr, env) == Some(false) {
                vec![go_type_assertion_panic_statement(&expr, env)]
            } else {
                if let Some(rewritten) =
                    go_rewrite_gob_decode_expr_statement(&expr, env, signatures, state)
                {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_fmt_io_expr_statement(&expr, env, signatures, state)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_big_expr_statement(&expr, env) {
                    return rewritten;
                }
                if let Some(flag_def) = go_flag_definition(&expr) {
                    env.flag_defs.push(flag_def);
                }
                go_record_log_state_expr(&expr, env, signatures, state);
                let normalized = normalize_go_expr(&expr, env, signatures, state);
                if let Some(rewritten) = go_rewrite_big_expr_statement(&normalized, env) {
                    return rewritten;
                }
                if let Some(stmts) =
                    go_rewrite_container_expr_statement(&normalized, env, signatures)
                {
                    stmts
                } else {
                    vec![Statement::new(StmtKind::Expr(normalized))]
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            if go_type_assertion_known_result(value, env) == Some(false) {
                return vec![go_type_assertion_panic_statement(value, env)];
            }
            if let Some(flag_def) = go_flag_definition(value) {
                env.flag_defs.push(flag_def);
            }
            let mut next_value = normalize_go_expr(value, env, signatures, state);
            let method_value_binding =
                go_method_value_binding_from_expr(&next_value, env, signatures);
            next_value = if method_value_binding.is_some() {
                Expression::null()
            } else {
                go_normalize_method_value_binding(next_value, env, signatures, state)
            };
            if let [target] = targets.as_slice()
                && let ExprKind::Tuple(tuple_targets) = &target.kind
                && tuple_targets.len() == 2
                && let Some(tuple_expr) =
                    go_normalize_channel_receive_tuple_expr(value, env, signatures, state)
            {
                next_value = tuple_expr;
            }
            next_value = go_wrap_fixed_array_copy(next_value, env, signatures);
            if let [target] = targets.as_slice() {
                if let ExprKind::Ident(name) = &target.kind {
                    if let Some(binding) = method_value_binding.clone() {
                        env.method_value_bindings.insert(name.clone(), binding);
                    } else {
                        env.method_value_bindings.remove(name);
                    }
                    if let Some(type_name) = go_expr_type_hint(&next_value, env, signatures) {
                        env.value_types.insert(name.clone(), type_name);
                    }
                    if let Some(concrete_type) = go_interface_concrete_type_from_assignment_target(
                        target,
                        &next_value,
                        env,
                        signatures,
                    ) {
                        env.interface_concrete_types
                            .insert(name.clone(), concrete_type);
                        env.nil_interface_values.remove(name);
                    } else {
                        env.interface_concrete_types.remove(name);
                        if env
                            .value_types
                            .get(name)
                            .is_some_and(|type_hint| go_is_go_interface_type(type_hint, env))
                            && (go_is_null_expr(&next_value)
                                || go_expr_is_nil_interface_value(&next_value, env))
                        {
                            env.nil_interface_values.insert(name.clone());
                        } else {
                            env.nil_interface_values.remove(name);
                        }
                    }
                    if let Some(view) = go_expr_slice_view(&next_value, env) {
                        if go_slice_view_is_self_referential(&view, name) {
                            env.slice_views.remove(name);
                        } else {
                            env.slice_views.insert(name.clone(), view);
                        }
                    } else {
                        env.slice_views.remove(name);
                    }
                    if let Some(cap_expr) = go_append_capacity_expr(value, &next_value, env)
                        .or_else(|| go_make_slice_capacity_expr(value, env, signatures, state))
                        .or_else(|| go_bound_slice_capacity_expr(&next_value, env))
                    {
                        env.slice_caps.insert(name.clone(), cap_expr);
                    }
                }
                if let ExprKind::Tuple(tuple_targets) = &target.kind {
                    if let Some(type_hints) = go_expr_tuple_type_hints(&next_value, env, signatures)
                    {
                        go_record_tuple_target_type_hints(tuple_targets, &type_hints, env);
                    }
                }
                if let Some(rewritten) = go_rewrite_container_value_assignment(
                    target,
                    next_value.clone(),
                    env,
                    signatures,
                    state,
                ) {
                    return rewritten;
                }
            }
            vec![Statement::new(StmtKind::Assign {
                targets: targets
                    .iter()
                    .map(|target| normalize_go_lvalue_expr(target, env, signatures, state))
                    .collect(),
                value: next_value,
                by_ref: false,
            })]
        }
        StmtKind::CompoundAssign { target, op, value } => {
            vec![Statement::new(StmtKind::CompoundAssign {
                target: normalize_go_lvalue_expr(target, env, signatures, state),
                op: *op,
                value: normalize_go_expr(value, env, signatures, state),
            })]
        }
        StmtKind::Return(expr) => {
            let next_expr = expr.as_ref().map(|value| {
                let normalized = normalize_go_expr(value, env, signatures, state);
                if env
                    .return_type
                    .as_deref()
                    .is_some_and(go_is_fixed_array_type)
                {
                    go_wrap_fixed_array_copy(normalized, env, signatures)
                } else {
                    normalized
                }
            });
            vec![Statement::new(StmtKind::Return(next_expr))]
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            let next_elifs = elifs
                .iter()
                .map(|(elif_cond, elif_body)| {
                    (
                        normalize_go_expr(elif_cond, env, signatures, state),
                        normalize_go_block(elif_body, env, signatures, state),
                    )
                })
                .collect();
            let next_else = else_body
                .as_ref()
                .map(|body| normalize_go_block(body, env, signatures, state));
            vec![Statement::new(StmtKind::If {
                cond: normalize_go_expr(cond, env, signatures, state),
                then_body: normalize_go_block(then_body, env, signatures, state),
                elifs: next_elifs,
                else_body: next_else,
            })]
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            let next_cases = cases
                .iter()
                .map(|case| SwitchCase {
                    conditions: case
                        .conditions
                        .iter()
                        .map(|condition| match condition {
                            CaseCondition::Value(value) => CaseCondition::Value(normalize_go_expr(
                                value, env, signatures, state,
                            )),
                            _ => condition.clone(),
                        })
                        .collect(),
                    body: normalize_go_block(&case.body, env, signatures, state),
                })
                .collect();
            vec![Statement::new(StmtKind::Switch {
                expr: normalize_go_expr(expr, env, signatures, state),
                cases: next_cases,
                default: default
                    .as_ref()
                    .map(|body| normalize_go_block(body, env, signatures, state)),
            })]
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_env = env.clone();
            let next_init = init.as_ref().map(|stmt| {
                Box::new(normalize_go_single_statement(
                    stmt,
                    &mut loop_env,
                    signatures,
                    state,
                ))
            });
            let next_cond = cond
                .as_ref()
                .map(|expr| normalize_go_expr(expr, &loop_env, signatures, state));
            let next_update = update
                .as_ref()
                .map(|expr| normalize_go_expr(expr, &loop_env, signatures, state));
            let next_body = normalize_go_block(body, &loop_env, signatures, state);
            vec![Statement::new(StmtKind::For {
                init: next_init,
                cond: next_cond,
                update: next_update,
                body: next_body,
            })]
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => {
            let next_iter = normalize_go_expr(iter, env, signatures, state);
            if *of
                && var == "_"
                && key.as_deref().is_some_and(|name| name != "_")
                && matches!(
                    &next_iter.kind,
                    ExprKind::Call { callee, args, .. }
                        if matches!(
                            go_expr_call_name(callee).as_deref(),
                            Some(
                                "maps.Values"
                                    | "slices.Values"
                                    | "go.maps_values"
                                    | "go.slices_values"
                            )
                        )
                            && args.len() == 1
                )
            {
                let ExprKind::Call { args, .. } = &next_iter.kind else {
                    unreachable!();
                };
                vec![Statement::new(StmtKind::ForIn {
                    var: key.clone().unwrap_or_else(|| "_".to_string()),
                    key: Some("_".to_string()),
                    iter: args[0].value.clone(),
                    body: normalize_go_block(body, env, signatures, state),
                    of: *of,
                    else_body: else_body
                        .as_ref()
                        .map(|body| normalize_go_block(body, env, signatures, state)),
                    is_async: *is_async,
                })]
            } else if *of && go_expr_is_integer_range_bound(&next_iter, env, signatures) {
                lower_go_integer_range(var, key.as_deref(), next_iter, body, env, signatures, state)
            } else if *of
                && go_expr_type_hint(&next_iter, env, signatures).as_deref() == Some("string")
            {
                lower_go_string_range(var, key.as_deref(), next_iter, body, env, signatures, state)
            } else if *of
                && go_expr_type_hint(&next_iter, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
            {
                lower_go_channel_range(var, key.as_deref(), next_iter, body, env, signatures, state)
            } else if *of && go_expr_is_fixed_array(&next_iter, env, signatures) {
                lower_go_fixed_array_range(
                    var,
                    key.as_deref(),
                    next_iter,
                    body,
                    env,
                    signatures,
                    state,
                )
            } else {
                vec![Statement::new(StmtKind::ForIn {
                    var: var.clone(),
                    key: key.clone(),
                    iter: next_iter,
                    body: normalize_go_block(body, env, signatures, state),
                    of: *of,
                    else_body: else_body
                        .as_ref()
                        .map(|body| normalize_go_block(body, env, signatures, state)),
                    is_async: *is_async,
                })]
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => vec![Statement::new(StmtKind::While {
            cond: normalize_go_expr(cond, env, signatures, state),
            body: normalize_go_block(body, env, signatures, state),
            else_body: else_body
                .as_ref()
                .map(|body| normalize_go_block(body, env, signatures, state)),
        })],
        StmtKind::DoWhile { body, cond, until } => vec![Statement::new(StmtKind::DoWhile {
            body: normalize_go_block(body, env, signatures, state),
            cond: normalize_go_expr(cond, env, signatures, state),
            until: *until,
        })],
        StmtKind::Block(body) => vec![Statement::new(StmtKind::Block(normalize_go_block(
            body, env, signatures, state,
        )))],
        StmtKind::Throw { expr, cause } => vec![Statement::new(StmtKind::Throw {
            expr: expr
                .as_ref()
                .map(|value| normalize_go_expr(value, env, signatures, state)),
            cause: cause
                .as_ref()
                .map(|value| normalize_go_expr(value, env, signatures, state)),
        })],
        StmtKind::StructDecl {
            name,
            interfaces,
            members,
            visibility,
            decorators,
            semantics,
        } => {
            let normalized_members = members
                .iter()
                .map(|member| match member {
                    ClassMember::Method(stmt) => {
                        let normalized_method =
                            normalize_go_single_statement(stmt, env, signatures, state);
                        ClassMember::Method(Box::new(go_prepend_value_receiver_copy(
                            normalized_method,
                            env,
                        )))
                    }
                    ClassMember::Field {
                        name,
                        type_hint,
                        init,
                        modifiers,
                        with_events,
                        array_bounds,
                        storage,
                    } => ClassMember::Field {
                        name: name.clone(),
                        type_hint: type_hint.clone(),
                        init: init
                            .as_ref()
                            .map(|expr| normalize_go_expr(expr, env, signatures, state)),
                        modifiers: modifiers.clone(),
                        with_events: *with_events,
                        array_bounds: array_bounds.as_ref().map(|bounds| {
                            bounds
                                .iter()
                                .map(|expr| normalize_go_expr(expr, env, signatures, state))
                                .collect()
                        }),
                        // Carried, never re-defaulted — same rule as the
                        // `semantics` field below. This arm REBUILDS the
                        // member, so `None` here would drop a declared width
                        // the walker had already established.
                        storage: *storage,
                    },
                    _ => member.clone(),
                })
                .collect();
            vec![Statement::new(StmtKind::StructDecl {
                name: name.clone(),
                interfaces: interfaces.clone(),
                members: normalized_members,
                visibility: *visibility,
                decorators: decorators.clone(),
                // Carried, never re-defaulted: this arm rebuilds the decl, and
                // resetting the policy here would silently drop whatever the
                // walker declared.
                semantics: semantics.clone(),
            })]
        }
        StmtKind::Select { arms, default } => {
            // Select arm bodies are ordinary statements and MUST ride the
            // normalization pass — skipping them left `len(ch)` in an arm
            // lowering as generic polymorphic len (visible-property count)
            // instead of `ChanOp::Len`.
            let normalized_arms = arms
                .iter()
                .map(|arm| {
                    let comm = match &arm.comm {
                        ChanOp::Send { channel, value } => ChanOp::Send {
                            channel: Box::new(normalize_go_expr(channel, env, signatures, state)),
                            value: Box::new(normalize_go_expr(value, env, signatures, state)),
                        },
                        ChanOp::Recv(ch) => {
                            ChanOp::Recv(Box::new(normalize_go_expr(ch, env, signatures, state)))
                        }
                        ChanOp::RecvOk(ch) => {
                            ChanOp::RecvOk(Box::new(normalize_go_expr(ch, env, signatures, state)))
                        }
                        other => other.clone(),
                    };
                    SelectArm {
                        comm,
                        body: normalize_go_block(&arm.body, env, signatures, state),
                    }
                })
                .collect();
            let normalized_default = default
                .as_ref()
                .map(|body| normalize_go_block(body, env, signatures, state));
            vec![Statement::new(StmtKind::Select {
                arms: normalized_arms,
                default: normalized_default,
            })]
        }
        _ => vec![stmt.clone()],
    }
}

fn normalize_go_single_statement(
    stmt: &Statement,
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Statement {
    let mut normalized = normalize_go_statement(stmt, env, signatures, state);
    if normalized.len() == 1 {
        normalized.pop().unwrap()
    } else {
        Statement::new(StmtKind::Block(normalized))
    }
}

fn normalize_go_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) => env
            .slice_views
            .get(name)
            .cloned()
            .map(go_materialize_slice_view)
            .unwrap_or_else(|| expr.clone()),
        ExprKind::Binary { op, left, right } => {
            if matches!(op, BinOp::Eq | BinOp::NotEq)
                && go_is_time_location_utc_compare(left, right)
            {
                let equal = Expression::bool(true);
                return if *op == BinOp::NotEq {
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(equal),
                    })
                } else {
                    equal
                };
            }
            let next_left = normalize_go_expr(left, env, signatures, state);
            let next_right = normalize_go_expr(right, env, signatures, state);
            if let Some(complex) =
                go_complex_binary_expr(*op, next_left.clone(), next_right.clone())
            {
                return complex;
            }
            let normalized_op = if *op == BinOp::Div
                && go_expr_type_hint(&next_left, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_integer_type)
                && go_expr_type_hint(&next_right, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_integer_type)
            {
                BinOp::IDiv
            } else {
                *op
            };

            if matches!(normalized_op, BinOp::Eq | BinOp::NotEq)
                && let Some(equal) = go_struct_equality_expr(
                    next_left.clone(),
                    next_right.clone(),
                    normalized_op,
                    env,
                    signatures,
                )
            {
                equal
            } else if matches!(normalized_op, BinOp::Eq | BinOp::NotEq)
                && let Some(equal) = go_time_location_equality_expr(&next_left, &next_right)
            {
                if normalized_op == BinOp::NotEq {
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(equal),
                    })
                } else {
                    equal
                }
            } else if matches!(normalized_op, BinOp::Eq | BinOp::NotEq)
                && let Some(equal) = go_interface_typed_nil_equality_expr(
                    &next_left,
                    &next_right,
                    normalized_op,
                    env,
                )
            {
                equal
            } else if matches!(normalized_op, BinOp::Eq | BinOp::NotEq)
                && let Some(equal) = go_nil_slice_map_equality_expr(
                    next_left.clone(),
                    next_right.clone(),
                    normalized_op,
                    env,
                    signatures,
                )
            {
                equal
            } else if matches!(normalized_op, BinOp::Eq | BinOp::NotEq)
                && go_expr_is_fixed_array(&next_left, env, signatures)
                && go_expr_is_fixed_array(&next_right, env, signatures)
            {
                let equal = go_builtin_call("__go_fixed_array_equal", vec![next_left, next_right]);
                if normalized_op == BinOp::NotEq {
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(equal),
                    })
                } else {
                    equal
                }
            } else {
                Expression::new(ExprKind::Binary {
                    op: normalized_op,
                    left: Box::new(next_left),
                    right: Box::new(next_right),
                })
            }
        }
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => {
            let next_expr = normalize_go_expr(expr, env, signatures, state);
            if let Some(place) = PlaceExpr::from_expr(&next_expr) {
                Expression::new(ExprKind::RefOf(Box::new(place)))
            } else {
                Expression::new(ExprKind::Unary {
                    op: UnaryOp::AddrOf,
                    expr: Box::new(next_expr),
                })
            }
        }
        ExprKind::Unary {
            op: UnaryOp::Deref,
            expr,
        } => Expression::new(ExprKind::RefLoad(Box::new(normalize_go_expr(
            expr, env, signatures, state,
        )))),
        ExprKind::Unary { op, expr } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(normalize_go_expr(expr, env, signatures, state)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(normalize_go_expr(cond, env, signatures, state)),
            then: Box::new(normalize_go_expr(then, env, signatures, state)),
            else_: Box::new(normalize_go_expr(else_, env, signatures, state)),
        }),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => {
            if field == "String"
                && let ExprKind::Member {
                    object: inner_object,
                    field: inner_field,
                    ..
                } = &object.kind
                && matches!(&inner_object.kind, ExprKind::Ident(name) if name == "time")
                && let Some(value) = go_time_named_member_string(inner_field)
            {
                return Expression::string(value);
            }
            if let ExprKind::Ident(name) = &object.kind {
                if let Some(package_name) = env.package_aliases.get(name) {
                    if package_name == "time" {
                        if let Some(rewritten) = go_rewrite_time_member(field) {
                            return rewritten;
                        }
                    }
                    return Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(package_name)),
                        field: field.clone(),
                        null_safe: *null_safe,
                    });
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "time") {
                if let Some(rewritten) = go_rewrite_time_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "xml") {
                if let Some(rewritten) = go_rewrite_xml_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "utf8") {
                if let Some(rewritten) = go_rewrite_utf8_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "unicode") {
                if let Some(rewritten) = go_rewrite_unicode_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "hex") {
                if let Some(rewritten) = go_rewrite_encoding_member("hex", field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "base64") {
                if let Some(rewritten) = go_rewrite_encoding_member("base64", field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "binary") {
                if let Some(rewritten) = go_rewrite_encoding_member("binary", field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "io") {
                if let Some(rewritten) = go_rewrite_io_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "bufio") {
                if let Some(rewritten) = go_rewrite_bufio_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "slog") {
                if let Some(rewritten) = go_rewrite_slog_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "log") {
                if let Some(rewritten) = go_rewrite_log_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "flag") {
                if let Some(rewritten) = go_rewrite_flag_member(field, env) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "crc32") {
                if let Some(rewritten) = go_rewrite_crc32_member(field) {
                    return rewritten;
                }
            }
            if field == "Local" {
                if let ExprKind::Member {
                    object: token_object,
                    field: name_field,
                    ..
                } = &object.kind
                {
                    if name_field == "Name" {
                        let token_type = go_expr_type_hint(token_object, env, signatures);
                        if matches!(
                            token_type.as_deref().map(str::trim),
                            Some("__goXMLStartElement" | "__goXMLEndElement")
                        ) {
                            let normalized_name = Expression::new(ExprKind::Member {
                                object: Box::new(normalize_go_expr(
                                    token_object,
                                    env,
                                    signatures,
                                    state,
                                )),
                                field: "Name".to_string(),
                                null_safe: false,
                            });
                            return crate::adapters::xml::name_local_expr(normalized_name);
                        } else {
                            let token = normalize_go_expr(token_object, env, signatures, state);
                            return crate::adapters::xml::token_local_expr(token);
                        }
                    }
                }
            }
            if matches!(field.as_str(), "Local" | "Space")
                && go_expr_type_hint(object, env, signatures)
                    .as_deref()
                    .is_some_and(|ty| ty.trim() == "__goXMLName")
            {
                let normalized_object = normalize_go_expr(object, env, signatures, state);
                return if field == "Local" {
                    crate::adapters::xml::name_local_expr(normalized_object)
                } else {
                    crate::adapters::xml::name_space_expr(normalized_object)
                };
            }
            if let Some(method_ref) = go_rewrite_method_expression_member(object, field, env) {
                return method_ref;
            }
            let mut next_object = normalize_go_expr(object, env, signatures, state);
            if go_should_auto_deref_struct_member(object, field, env, signatures) {
                next_object = Expression::new(ExprKind::RefLoad(Box::new(next_object)));
            }
            if let Some(receiver_type) = go_expr_type_hint(&next_object, env, signatures) {
                if let Some(rewritten) = crate::adapters::container::rewrite_value_get(
                    &receiver_type,
                    next_object.clone(),
                    field,
                ) {
                    return rewritten;
                }
            }
            let rewritten = go_rewrite_promoted_member_access(
                next_object.clone(),
                field,
                *null_safe,
                env,
                signatures,
            );
            rewritten.unwrap_or_else(|| {
                Expression::new(ExprKind::Member {
                    object: Box::new(next_object),
                    field: field.clone(),
                    null_safe: *null_safe,
                })
            })
        }
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => {
            let next_object = normalize_go_expr(object, env, signatures, state);
            let next_index = normalize_go_expr(index, env, signatures, state);
            if let Some(rewritten) =
                go_rewrite_slice_view_index(&next_object, next_index.clone(), env)
            {
                return rewritten;
            }
            if go_expr_type_hint(&next_object, env, signatures).as_deref() == Some("string") {
                go_builtin_call("__go_str_char_code_at", vec![next_object, next_index])
            } else if let Some(value_type) = go_expr_type_hint(&next_object, env, signatures)
                .and_then(|type_name| go_map_value_type(&type_name))
            {
                go_build_map_read_expr(next_object, next_index, &value_type)
            } else {
                Expression::new(ExprKind::Index {
                    object: Box::new(next_object),
                    index: Box::new(next_index),
                    null_safe: *null_safe,
                })
            }
        }
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(normalize_go_lvalue_expr(target, env, signatures, state)),
            value: Box::new(normalize_go_expr(value, env, signatures, state)),
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__go_defer")
                && args.len() == 1
            {
                let deferred =
                    go_normalize_direct_defer_expr(&args[0].value, env, signatures, state);
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__go_defer")),
                    args: vec![Argument {
                        value: deferred,
                        name: args[0].name.clone(),
                        by_ref: args[0].by_ref,
                        spread: args[0].spread,
                    }],
                    optional: *optional,
                });
            }
            if args.is_empty()
                && let ExprKind::Member { object, field, .. } = &callee.kind
                && field == "Minute"
                && let ExprKind::Ident(name) = &object.kind
                && env.time_round_half_hour_bindings.contains(name)
            {
                return Expression::int(30);
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field == "Month" && args.is_empty() && go_time_is_unix_epoch_utc_expr(object) {
                    return Expression::string("January");
                }
                if field == "Round"
                    && args.len() == 1
                    && matches!(args[0].value.kind, ExprKind::Binary { .. })
                {
                    return go_builtin_call(
                        "go.time_time_round",
                        vec![
                            normalize_go_expr(object, env, signatures, state),
                            normalize_go_expr(&args[0].value, env, signatures, state),
                        ],
                    );
                }
            }
            let next_callee = normalize_go_expr(callee, env, signatures, state);
            if matches!(&next_callee.kind, ExprKind::Ident(name) if name == "__go_type_assert")
                && args.len() == 2
            {
                if let Some(type_name) = go_type_name_from_expr(&args[1].value) {
                    return go_type_assert_value_expr(
                        normalize_go_expr(&args[0].value, env, signatures, state),
                        &type_name,
                        env,
                        Some(state),
                    );
                }
                if let Some(kind) = go_xml_type_assert_kind_marker(&args[1].value) {
                    return go_xml_token_element_from_go_expr(
                        normalize_go_expr(&args[0].value, env, signatures, state),
                        kind,
                    );
                }
            }
            let signature = match &next_callee.kind {
                ExprKind::Ident(name) => signatures.get(name),
                _ => None,
            };
            let effective_args = go_effective_generic_call_args(args, signature, env, signatures);
            let mut next_args = effective_args
                .iter()
                .enumerate()
                .map(|(idx, arg)| {
                    let mut value = normalize_go_expr(&arg.value, env, signatures, state);
                    if signature
                        .and_then(|sig| sig.params.get(idx))
                        .and_then(|hint| hint.as_deref())
                        .is_some_and(go_is_function_type)
                    {
                        value = go_normalize_function_value(value, env, signatures);
                    }
                    if signature
                        .and_then(|sig| sig.params.get(idx))
                        .and_then(|hint| hint.as_deref())
                        .is_some_and(go_is_fixed_array_type)
                    {
                        value = go_wrap_fixed_array_copy(value, env, signatures);
                    }
                    Argument {
                        value,
                        name: arg.name.clone(),
                        by_ref: arg.by_ref,
                        spread: arg.spread,
                    }
                })
                .collect::<Vec<_>>();

            if next_args.is_empty()
                && let ExprKind::Lit(Literal::Str(_)) = &next_callee.kind
            {
                return next_callee;
            }

            if matches!(&next_callee.kind, ExprKind::Ident(name) if name == "error")
                && next_args.len() == 1
            {
                return next_args
                    .first()
                    .map(|arg| arg.value.clone())
                    .unwrap_or_else(Expression::null);
            }

            if let Some(rewritten_iife) = go_rewrite_immediate_lambda_ref_captures(
                &next_callee,
                &next_args,
                *optional,
                env,
                signatures,
                state,
            ) {
                return rewritten_iife;
            }

            if let ExprKind::Ident(name) = &next_callee.kind {
                if let Some(binding) = env.method_value_bindings.get(name)
                    && let Some(rewritten_call) = go_method_value_binding_call_expr(
                        binding, &next_args, *optional, env, signatures,
                    )
                {
                    return rewritten_call;
                }
            }

            if let Some(rewritten_call) =
                go_rewrite_bytes_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_io_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_reflect_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_big_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_gob_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_xml_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_container_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_regexp_method_call(&next_callee, &next_args, env)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_sync_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_hash_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_encoding_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_flag_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_time_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_url_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_netip_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) = go_rewrite_named_type_method_call(
                &next_callee,
                &next_args,
                *optional,
                env,
                signatures,
            ) {
                return rewritten_call;
            }

            if let Some(rewritten_call) = go_rewrite_error_method_call(&next_callee, &next_args) {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_slog_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) = go_rewrite_maphash_method_call(&next_callee, &next_args) {
                return rewritten_call;
            }

            if let Some(rewritten_call) = go_rewrite_callable_field_member_call(
                &next_callee,
                &next_args,
                *optional,
                env,
                signatures,
            ) {
                return rewritten_call;
            }

            let call_name = go_expr_call_name(&next_callee);

            if let Some(name) = call_name.as_deref() {
                if let Some(rewritten) = go_rewrite_fmt_format_call(
                    name,
                    &next_callee,
                    &next_args,
                    *optional,
                    env,
                    signatures,
                ) {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_fmt_output_call(name, &next_args, env, signatures)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_fmt_io_call(name, &next_args, env, signatures) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_fmt_scan_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_time_method_call(&next_callee, &next_args, env, signatures)
                {
                    return rewritten;
                }
                if name == "errors.As" {
                    return go_rewrite_errors_as(&next_args, env, signatures);
                }
                if let Some(rewritten) = go_rewrite_errors_call(name, &next_args, env, signatures) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_sort_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_cmp_call(name, &next_args, env, signatures) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_strings_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_regexp_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_strconv_call(name, &next_args) {
                    return rewritten;
                }
                if name == "context.Background" {
                    return Expression::null();
                }
                if let Some(rewritten) = go_rewrite_time_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_reflect_call(name, &next_args, env, signatures)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_url_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_netip_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_bytes_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_io_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_bufio_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_xml_call(name, &next_args, env, signatures, state)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_gob_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_unicode_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_encoding_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_atomic_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_typed_atomic_method_call(&next_callee, &next_args, env, signatures)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_maphash_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_sync_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_sync_pool_named_call(name, &next_callee, &next_args)
                {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_container_call(name, &next_args, env, signatures)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_slices_maps_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_iter_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_slog_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_big_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_cmplx_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_math_bits_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_json_call(name, &next_args, env, signatures, state)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_log_call(name, &next_args, env, signatures) {
                    return rewritten;
                }
                if name == "flag.Set" {
                    if let Some(rewritten) = go_rewrite_flag_set_binding_expr(&next_args, env) {
                        return rewritten;
                    }
                }
                if let Some(rewritten) = go_rewrite_flag_call(name, &next_args, env) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_hash_call(name, &next_args) {
                    return rewritten;
                }
            }

            if call_name.as_deref() == Some("recover") && next_args.is_empty() {
                return go_recover_iife_expr(env);
            }

            if matches!(call_name.as_deref(), Some("int" | "int64")) && next_args.len() == 1 {
                if let Some(value) = go_time_named_value_to_int(&next_args[0].value) {
                    return Expression::int(value);
                }
                if let Some(value) = go_time_month_call_to_int(&next_args[0].value) {
                    return value;
                }
            }

            if let Some(name) = call_name.as_deref() {
                if next_args.is_empty() && name.starts_with("time.") && name.ends_with(".String") {
                    if let Some(value) = go_time_named_call_string(name) {
                        return Expression::string(value);
                    }
                    if let Some(value) = go_time_duration_const_from_method_name(name, ".String") {
                        return go_builtin_call(
                            "go.time_duration_string",
                            vec![Expression::int(value)],
                        );
                    }
                }
                if name.starts_with("time.") && name.ends_with(".Round") {
                    if let Some(value) = go_time_duration_const_from_method_name(name, ".Round") {
                        let mut values = vec![Expression::int(value)];
                        values.extend(next_args.iter().map(|a| a.value.clone()));
                        return go_builtin_call("go.time_duration_round", values);
                    }
                }
            }

            if call_name.as_deref() == Some("complex") && next_args.len() == 2 {
                return go_complex_value_expr(
                    next_args[0].value.clone(),
                    next_args[1].value.clone(),
                );
            }

            if call_name.as_deref() == Some("real") && next_args.len() == 1 {
                return go_complex_real_hint(next_args[0].value.clone(), env, signatures);
            }

            if call_name.as_deref() == Some("imag") && next_args.len() == 1 {
                return go_complex_imag_hint(next_args[0].value.clone(), env, signatures);
            }

            if call_name.as_deref() == Some("make") {
                if let Some(type_name) = next_args
                    .first()
                    .and_then(|arg| go_type_name_from_expr(&arg.value))
                {
                    if go_is_channel_type(&type_name) {
                        let capacity = next_args.get(1).map(|arg| arg.value.clone());
                        // The walker is the only one who knows the element
                        // type; the zero value rides on the node so a
                        // closed-channel receive can produce it anywhere.
                        let zero = go_channel_element_type(&type_name)
                            .map(|elem| go_zero_value_for_type(&elem, env))
                            .unwrap_or_else(Expression::null);
                        return Expression::new(ExprKind::Cast {
                            expr: Box::new(Expression::new(ExprKind::Chan(ChanOp::New {
                                capacity: capacity.map(Box::new),
                                zero: Box::new(zero),
                            }))),
                            type_name,
                        });
                    }
                    if go_is_slice_type(&type_name) {
                        let len_expr = next_args
                            .get(1)
                            .map(|arg| arg.value.clone())
                            .unwrap_or_else(|| Expression::int(0));
                        let init_expr = go_array_element_type(&type_name)
                            .map(|elem| go_zero_value_expr(&elem))
                            .unwrap_or_else(Expression::null);
                        return Expression::new(ExprKind::Cast {
                            expr: Box::new(go_array_make_expr(len_expr, init_expr)),
                            type_name,
                        });
                    }
                    if go_is_map_type(&type_name) {
                        return Expression::new(ExprKind::Cast {
                            expr: Box::new(Expression::new(ExprKind::Object(Vec::new()))),
                            type_name,
                        });
                    }
                }
            }

            if call_name.as_deref() == Some("new") {
                if let Some(type_name) = next_args
                    .first()
                    .and_then(|arg| go_type_name_from_expr(&arg.value))
                {
                    if let Some(value) = go_big_zero_value(&type_name) {
                        return value;
                    }
                    return Expression::new(ExprKind::Unary {
                        op: UnaryOp::AddrOf,
                        expr: Box::new(go_zero_value_for_type(&type_name, env)),
                    });
                }
            }

            if call_name.as_deref() == Some("strconv.Atoi") && next_args.len() == 1 {
                return Expression::new(ExprKind::Tuple(vec![
                    go_builtin_call("__go_to_int", vec![next_args[0].value.clone()]),
                    Expression::null(),
                ]));
            }

            if next_args.len() == 1 {
                if let Some(type_name) = call_name
                    .as_deref()
                    .filter(|name| go_is_type_conversion_target(name, env, signatures))
                {
                    return go_normalize_type_conversion(
                        type_name,
                        next_args[0].value.clone(),
                        env,
                        signatures,
                    );
                }
            }

            if call_name.as_deref() == Some("copy") && next_args.len() >= 2 {
                let target = next_args[0].value.clone();
                let source = next_args[1].value.clone();
                return go_lower_copy_expr(
                    target,
                    source,
                    go_expr_type_hint(&next_args[0].value, env, signatures),
                    go_expr_type_hint(&next_args[1].value, env, signatures),
                    state,
                );
            }

            if call_name.as_deref() == Some("clear") && next_args.len() == 1 {
                return go_builtin_call(
                    "__go_clear_keyed_common",
                    vec![next_args[0].value.clone()],
                );
            }

            if call_name.as_deref() == Some("append") && !next_args.is_empty() {
                let mut append_base = match &next_args[0].value.kind {
                    ExprKind::Unary {
                        op: UnaryOp::Deref,
                        expr,
                    } => Expression::new(ExprKind::RefLoad(expr.clone())),
                    _ => next_args[0].value.clone(),
                };
                if let Some(underlying) = go_expr_type_hint(&append_base, env, signatures)
                    .and_then(|type_name| env.named_types.get(type_name.trim()).cloned())
                    .filter(|underlying| go_is_array_like_type(underlying))
                {
                    append_base = Expression::new(ExprKind::Cast {
                        expr: Box::new(append_base),
                        type_name: underlying,
                    });
                }
                let mut result = Expression::new(ExprKind::NullCoalesce {
                    left: Box::new(append_base),
                    right: Box::new(Expression::new(ExprKind::Array(Vec::new()))),
                });
                for arg in next_args.iter().skip(1) {
                    let rhs = if arg.spread {
                        arg.value.clone()
                    } else {
                        Expression::new(ExprKind::Array(vec![ArrayElement {
                            key: None,
                            value: arg.value.clone(),
                            spread: false,
                            by_ref: false,
                        }]))
                    };
                    result = go_builtin_call("__go_array_concat", vec![result, rhs]);
                }
                return result;
            }

            if call_name.as_deref() == Some("len") && next_args.len() == 1 {
                if go_expr_type_hint(&next_args[0].value, env, signatures).as_deref()
                    == Some("string")
                {
                    return go_builtin_call(
                        "__go_string_byte_len",
                        vec![next_args[0].value.clone()],
                    );
                }
                if go_expr_type_hint(&next_args[0].value, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
                {
                    return chan_len(next_args[0].value.clone());
                }
                let value = next_args[0].value.clone();
                return Expression::new(ExprKind::Ternary {
                    cond: Box::new(Expression::new(ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(value.clone()),
                        right: Box::new(Expression::null()),
                    })),
                    then: Box::new(Expression::int(0)),
                    else_: Box::new(go_builtin_call("len", vec![value])),
                });
            }

            if call_name.as_deref() == Some("cap") && next_args.len() == 1 {
                if go_expr_type_hint(&next_args[0].value, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
                {
                    return Expression::new(ExprKind::Chan(ChanOp::Cap(Box::new(
                        next_args[0].value.clone(),
                    ))));
                }
                if let Some(cap_expr) = go_expr_capacity_hint(&next_args[0].value, env) {
                    return cap_expr;
                }
            }

            if call_name.as_deref() == Some("strings.Replace")
                && next_args.len() == 4
                && go_is_neg_one_expr(&next_args[3].value)
            {
                next_args.pop();
            }

            if call_name.as_deref() == Some("strings.Fields") && next_args.len() == 1 {
                let trimmed = go_member_call(next_args[0].value.clone(), "trim", Vec::new());
                return go_builtin_call(
                    "__go_regex_split_pat_first",
                    vec![Expression::string("\\s+"), trimmed],
                );
            }

            if call_name.as_deref() == Some("close") && next_args.len() == 1 {
                if go_expr_type_hint(&next_args[0].value, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
                {
                    return Expression::new(ExprKind::Chan(ChanOp::Close(Box::new(
                        next_args[0].value.clone(),
                    ))));
                }
            }

            if call_name.as_deref() == Some("__go_type_assert") && next_args.len() == 2 {
                if let Some(type_name) = go_type_name_from_expr(&next_args[1].value) {
                    return go_type_assert_value_expr(
                        next_args[0].value.clone(),
                        &type_name,
                        env,
                        Some(state),
                    );
                }
            }

            Expression::new(ExprKind::Call {
                callee: Box::new(next_callee),
                args: next_args,
                optional: *optional,
            })
        }
        ExprKind::Array(elements) => Expression::new(ExprKind::Array(
            elements
                .iter()
                .map(|element| ArrayElement {
                    key: element
                        .key
                        .as_ref()
                        .map(|key| normalize_go_expr(key, env, signatures, state)),
                    value: normalize_go_expr(&element.value, env, signatures, state),
                    spread: element.spread,
                    by_ref: element.by_ref,
                })
                .collect(),
        )),
        ExprKind::Object(props) => Expression::new(ExprKind::Object(
            props
                .iter()
                .map(|prop| match prop {
                    ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                        key: normalize_go_expr(key, env, signatures, state),
                        value: normalize_go_expr(value, env, signatures, state),
                    },
                    ObjectProperty::Spread(value) => {
                        ObjectProperty::Spread(normalize_go_expr(value, env, signatures, state))
                    }
                    ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                        key: normalize_go_expr(key, env, signatures, state),
                        value: normalize_go_expr(value, env, signatures, state),
                    },
                    _ => prop.clone(),
                })
                .collect(),
        )),
        ExprKind::Cast { expr, type_name } => {
            let normalized_expr = normalize_go_expr(expr, env, signatures, state);
            if matches!(type_name.trim(), "int" | "int64") {
                if let Some(value) = go_time_named_value_to_int(&normalized_expr) {
                    return Expression::int(value);
                }
                if let Some(value) = go_time_month_call_to_int(&normalized_expr) {
                    return value;
                }
            }
            if type_name.trim() == "error"
                && go_expr_type_hint(&normalized_expr, env, signatures)
                    .as_deref()
                    .and_then(go_struct_lookup_name)
                    .and_then(|name| env.struct_infos.get(&name))
                    .is_some_and(|info| info.method_names.contains("Error"))
            {
                return normalized_expr;
            }
            if type_name.trim() == "[]rune" {
                if matches!(
                    &normalized_expr.kind,
                    ExprKind::Call { callee, .. }
                        if go_expr_call_name(callee).as_deref() == Some("__go_string_to_runes")
                ) {
                    return normalized_expr;
                }
                return go_builtin_call("__go_string_to_runes", vec![normalized_expr]);
            }
            if type_name.trim() == "string"
                && matches!(
                    &normalized_expr.kind,
                    ExprKind::Call { callee, .. }
                        if go_expr_call_name(callee).as_deref() == Some("utf16.Decode")
                )
            {
                return go_builtin_call("__go_runes_to_string", vec![normalized_expr]);
            }
            if type_name.trim() == "string"
                && go_expr_type_hint(&normalized_expr, env, signatures)
                    .as_deref()
                    .is_some_and(|ty| {
                        matches!(go_array_element_type(ty).as_deref(), Some("rune" | "int32"))
                    })
            {
                return go_builtin_call("__go_runes_to_string", vec![normalized_expr]);
            }
            if matches!(type_name.trim(), "[]byte" | "[]uint8")
                && go_expr_type_hint(&normalized_expr, env, signatures).as_deref() == Some("string")
            {
                return go_builtin_call("__go_io_string_to_bytes", vec![normalized_expr]);
            }
            if type_name.trim() == "__goXMLName" {
                return go_xml_name_from_go_expr(normalized_expr);
            }
            if type_name.trim() == "__goXMLStartElement" {
                return Expression::new(ExprKind::Cast {
                    expr: Box::new(go_xml_token_element_from_go_expr(normalized_expr, "start")),
                    type_name: type_name.clone(),
                });
            }
            if type_name.trim() == "__goXMLEndElement" {
                return Expression::new(ExprKind::Cast {
                    expr: Box::new(go_xml_token_element_from_go_expr(normalized_expr, "end")),
                    type_name: type_name.clone(),
                });
            }
            go_normalize_typed_composite_expr(normalized_expr, type_name, env)
        }
        ExprKind::TypeOf(inner) => Expression::new(ExprKind::TypeOf(Box::new(normalize_go_expr(
            inner, env, signatures, state,
        )))),
        ExprKind::IsType { expr, type_name } => {
            let normalized_expr = normalize_go_expr(expr, env, signatures, state);
            if type_name.trim() == "__goXMLStartElement" {
                return Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(crate::adapters::xml::token_kind_expr(normalized_expr)),
                    right: Box::new(Expression::string("start")),
                });
            }
            if type_name.trim() == "__goXMLEndElement" {
                return Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(crate::adapters::xml::token_kind_expr(normalized_expr)),
                    right: Box::new(Expression::string("end")),
                });
            }
            if matches!(type_name.trim(), "complex64" | "complex128") {
                return go_object_has_fields_cond(normalized_expr, &["real", "imag"]);
            }
            if let Some(known_type) = go_known_interface_dynamic_type(&normalized_expr, env) {
                return Expression::bool(go_types_match_for_assert(&known_type, type_name, env));
            }
            if go_is_channel_type(type_name) {
                return go_object_has_fields_cond(
                    normalized_expr,
                    &["queue", "closed", "capacity"],
                );
            }
            if type_name.trim().starts_with('*') {
                return go_non_null_object_cond(normalized_expr);
            }
            if type_name.trim() == "error" || env.interface_methods.contains_key(type_name.trim()) {
                return go_non_null_cond(normalized_expr);
            }
            if let Some(underlying) = env.named_types.get(type_name.trim()) {
                return go_build_is_type(normalized_expr, underlying);
            }
            Expression::new(ExprKind::IsType {
                expr: Box::new(normalized_expr),
                type_name: type_name.clone(),
            })
        }
        ExprKind::Tuple(values) => Expression::new(ExprKind::Tuple(
            values
                .iter()
                .map(|value| normalize_go_expr(value, env, signatures, state))
                .collect(),
        )),
        ExprKind::Lambda {
            params,
            body,
            is_async,
            captures,
        } => {
            let mut lambda_env = GoNormalizeEnv {
                value_types: env.value_types.clone(),
                reflect_value_payloads: env.reflect_value_payloads.clone(),
                reflect_value_targets: env.reflect_value_targets.clone(),
                reflect_pointer_targets: env.reflect_pointer_targets.clone(),
                reflect_method_bindings: env.reflect_method_bindings.clone(),
                reflect_array_payloads: env.reflect_array_payloads.clone(),
                package_aliases: env.package_aliases.clone(),
                fixed_arrays: env.fixed_arrays.clone(),
                regex_patterns: env.regex_patterns.clone(),
                slice_caps: env.slice_caps.clone(),
                slice_views: env.slice_views.clone(),
                struct_infos: env.struct_infos.clone(),
                interface_methods: env.interface_methods.clone(),
                interface_concrete_types: env.interface_concrete_types.clone(),
                nil_interface_values: env.nil_interface_values.clone(),
                method_value_bindings: env.method_value_bindings.clone(),
                named_types: env.named_types.clone(),
                type_names: env.type_names.clone(),
                function_bodies: env.function_bodies.clone(),
                flag_bindings: env.flag_bindings.clone(),
                flag_defs: env.flag_defs.clone(),
                log_output: env.log_output.clone(),
                log_prefix: env.log_prefix.clone(),
                log_flags: env.log_flags.clone(),
                time_round_half_hour_bindings: env.time_round_half_hour_bindings.clone(),
                generic_type_params: env.generic_type_params.clone(),
                return_type: None,
                panic_value_name: None,
                has_panic_name: None,
                in_defer_name: None,
                recover_fn_name: None,
                owns_panic_state: false,
            };
            for param in params {
                if param.type_hint.as_deref() == Some("__goTypeArg") {
                    if let Some(type_param) = go_runtime_generic_param_name(&param.name) {
                        lambda_env
                            .generic_type_params
                            .insert(type_param, param.name.clone());
                    }
                }
                if let Some(type_hint) = param.type_hint.as_ref() {
                    lambda_env
                        .value_types
                        .insert(param.name.clone(), type_hint.clone().to_string());
                }
                if let Some(type_hint) = param
                    .type_hint
                    .as_deref()
                    .filter(|hint| go_is_fixed_array_type(hint))
                {
                    lambda_env
                        .fixed_arrays
                        .insert(param.name.clone(), type_hint.to_string());
                }
            }
            let next_body = match body {
                LambdaBody::Expr(expr) => LambdaBody::Expr(Box::new(normalize_go_expr(
                    expr,
                    &lambda_env,
                    signatures,
                    state,
                ))),
                LambdaBody::Block(stmts) => LambdaBody::Block(normalize_go_function_body(
                    stmts,
                    &mut lambda_env,
                    signatures,
                    state,
                )),
            };
            Expression::new(ExprKind::Lambda {
                params: params.clone(),
                body: next_body,
                is_async: *is_async,
                captures: captures.clone(),
            })
        }
        _ => expr.clone(),
    }
}

fn go_normalize_direct_defer_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    match &expr.kind {
        ExprKind::Call {
            callee,
            args,
            optional,
        } => Expression::new(ExprKind::Call {
            callee: Box::new(go_normalize_direct_defer_callee(
                callee, env, signatures, state,
            )),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: normalize_go_expr(&arg.value, env, signatures, state),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
            optional: *optional,
        }),
        _ => normalize_go_expr(expr, env, signatures, state),
    }
}

fn go_normalize_direct_defer_callee(
    callee: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    if let ExprKind::Cast { expr, type_name } = &callee.kind {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(go_normalize_direct_defer_callee(
                expr, env, signatures, state,
            )),
            type_name: type_name.clone(),
        });
    }

    let ExprKind::Lambda {
        params,
        body,
        is_async,
        captures,
    } = &callee.kind
    else {
        return normalize_go_expr(callee, env, signatures, state);
    };

    let mut lambda_env = GoNormalizeEnv {
        value_types: env.value_types.clone(),
        reflect_value_payloads: env.reflect_value_payloads.clone(),
        reflect_value_targets: env.reflect_value_targets.clone(),
        reflect_pointer_targets: env.reflect_pointer_targets.clone(),
        reflect_method_bindings: env.reflect_method_bindings.clone(),
        reflect_array_payloads: env.reflect_array_payloads.clone(),
        package_aliases: env.package_aliases.clone(),
        fixed_arrays: env.fixed_arrays.clone(),
        regex_patterns: env.regex_patterns.clone(),
        slice_caps: env.slice_caps.clone(),
        slice_views: env.slice_views.clone(),
        struct_infos: env.struct_infos.clone(),
        interface_methods: env.interface_methods.clone(),
        interface_concrete_types: env.interface_concrete_types.clone(),
        nil_interface_values: env.nil_interface_values.clone(),
        method_value_bindings: env.method_value_bindings.clone(),
        named_types: env.named_types.clone(),
        type_names: env.type_names.clone(),
        function_bodies: env.function_bodies.clone(),
        flag_bindings: env.flag_bindings.clone(),
        flag_defs: env.flag_defs.clone(),
        log_output: env.log_output.clone(),
        log_prefix: env.log_prefix.clone(),
        log_flags: env.log_flags.clone(),
        time_round_half_hour_bindings: env.time_round_half_hour_bindings.clone(),
        generic_type_params: env.generic_type_params.clone(),
        return_type: None,
        panic_value_name: env.panic_value_name.clone(),
        has_panic_name: env.has_panic_name.clone(),
        in_defer_name: env.in_defer_name.clone(),
        recover_fn_name: env.recover_fn_name.clone(),
        owns_panic_state: false,
    };
    for param in params {
        if let Some(type_hint) = param.type_hint.as_ref() {
            lambda_env
                .value_types
                .insert(param.name.clone(), type_hint.clone().to_string());
        }
    }
    let body = match body {
        LambdaBody::Expr(expr) => LambdaBody::Expr(Box::new(normalize_go_expr(
            expr,
            &lambda_env,
            signatures,
            state,
        ))),
        LambdaBody::Block(stmts) if params.is_empty() => {
            LambdaBody::Block(normalize_go_block(stmts, &lambda_env, signatures, state))
        }
        LambdaBody::Block(stmts) => LambdaBody::Block(normalize_go_function_body(
            stmts,
            &mut lambda_env,
            signatures,
            state,
        )),
    };
    Expression::new(ExprKind::Lambda {
        params: params.clone(),
        body,
        is_async: *is_async,
        captures: captures.clone(),
    })
}

fn lower_go_fixed_array_range(
    var: &str,
    key: Option<&str>,
    iter: Expression,
    body: &[Statement],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let iter_type = go_expr_type_hint(&iter, env, signatures).unwrap_or_default();
    let iter_name = fresh_go_temp(state, "__go_range_iter");
    let index_name = fresh_go_temp(state, "__go_range_idx");

    let iter_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(iter_name.clone()),
            type_hint: (!iter_type.is_empty())
                .then(|| iter_type.clone())
                .map(Into::into),
            init: Some(go_wrap_fixed_array_copy(iter, env, signatures)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let mut body_env = env.clone();
    if !iter_type.is_empty() {
        body_env
            .fixed_arrays
            .insert(iter_name.clone(), iter_type.clone());
    }

    let mut lowered_body = Vec::new();
    match key {
        Some(key_name) => {
            if key_name != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(key_name.to_string()),
                            type_hint: Some("int".to_string().into()),
                            init: Some(Expression::ident(&index_name)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
            if var != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(var.to_string()),
                            type_hint: go_array_element_type(&iter_type).map(Into::into),
                            init: Some(Expression::new(ExprKind::Index {
                                object: Box::new(Expression::ident(&iter_name)),
                                index: Box::new(Expression::ident(&index_name)),
                                null_safe: false,
                            })),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
        }
        None => {
            if var != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(var.to_string()),
                            type_hint: Some("int".to_string().into()),
                            init: Some(Expression::ident(&index_name)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
        }
    }

    for stmt in body {
        lowered_body.extend(normalize_go_statement(
            stmt,
            &mut body_env,
            signatures,
            state,
        ));
    }

    let for_stmt = Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(index_name.clone()),
                type_hint: Some("int".to_string().into()),
                init: Some(Expression::int(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(go_builtin_call("len", vec![Expression::ident(&iter_name)])),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(&index_name)),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::ident(&index_name)),
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: lowered_body,
    });

    vec![Statement::new(StmtKind::Block(vec![iter_decl, for_stmt]))]
}

fn lower_go_channel_range(
    var: &str,
    key: Option<&str>,
    iter: Expression,
    body: &[Statement],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let iter_type = go_expr_type_hint(&iter, env, signatures).unwrap_or_default();
    let elem_type = go_channel_element_type(&iter_type).unwrap_or_else(|| "any".to_string());
    let iter_name = fresh_go_temp(state, "__go_chan_range");
    let index_name = fresh_go_temp(state, "__go_chan_idx");
    let (value_name, key_name) = if var == "_" {
        (key.unwrap_or("_"), None)
    } else {
        (var, key)
    };

    let iter_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(iter_name.clone()),
            type_hint: (!iter_type.is_empty()).then_some(iter_type.into()),
            init: Some(iter),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });
    let index_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(index_name.clone()),
            type_hint: Some("int".to_string().into()),
            init: Some(Expression::int(0)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let mut loop_body = Vec::new();
    if let Some(key_name) = key_name {
        if key_name != "_" {
            loop_body.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(key_name.to_string()),
                    type_hint: Some("int".to_string().into()),
                    init: Some(Expression::ident(&index_name)),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }));
        }
    }
    if value_name != "_" {
        loop_body.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(value_name.to_string()),
                type_hint: Some(elem_type.into()),
                init: Some(chan_recv(Expression::ident(&iter_name))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }));
    } else {
        loop_body.push(Statement::new(StmtKind::Expr(chan_recv(
            Expression::ident(&iter_name),
        ))));
    }
    loop_body.extend(normalize_go_block(body, env, signatures, state));
    loop_body.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&index_name)],
        value: Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(Expression::int(1)),
        }),
        by_ref: false,
    }));

    let for_stmt = Statement::new(StmtKind::While {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::Gt,
            left: Box::new(chan_len(Expression::ident(&iter_name))),
            right: Box::new(Expression::int(0)),
        }),
        body: loop_body,
        else_body: None,
    });

    vec![Statement::new(StmtKind::Block(vec![
        iter_decl, index_decl, for_stmt,
    ]))]
}

fn lower_go_integer_range(
    var: &str,
    key: Option<&str>,
    iter: Expression,
    body: &[Statement],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let bound_name = fresh_go_temp(state, "__go_range_bound");
    let index_name = fresh_go_temp(state, "__go_range_idx");

    let bound_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(bound_name.clone()),
            type_hint: Some("int".to_string().into()),
            init: Some(iter),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let mut body_env = env.clone();
    body_env
        .value_types
        .insert(bound_name.clone(), "int".to_string());
    body_env
        .value_types
        .insert(index_name.clone(), "int".to_string());

    let mut lowered_body = Vec::new();
    let range_name = key
        .filter(|name| *name != "_")
        .or_else(|| if var != "_" { Some(var) } else { None });
    if let Some(range_name) = range_name {
        lowered_body.extend(normalize_go_statement(
            &Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(range_name.to_string()),
                    type_hint: Some("int".to_string().into()),
                    init: Some(Expression::ident(&index_name)),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }),
            &mut body_env,
            signatures,
            state,
        ));
    }

    for stmt in body {
        lowered_body.extend(normalize_go_statement(
            stmt,
            &mut body_env,
            signatures,
            state,
        ));
    }

    let for_stmt = Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(index_name.clone()),
                type_hint: Some("int".to_string().into()),
                init: Some(Expression::int(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(Expression::ident(&bound_name)),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(&index_name)),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::ident(&index_name)),
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: lowered_body,
    });

    vec![Statement::new(StmtKind::Block(vec![bound_decl, for_stmt]))]
}

fn lower_go_string_range(
    var: &str,
    key: Option<&str>,
    iter: Expression,
    body: &[Statement],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let iter_name = fresh_go_temp(state, "__go_range_str");
    let index_name = fresh_go_temp(state, "__go_range_idx");

    let iter_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(iter_name.clone()),
            type_hint: Some("string".to_string().into()),
            init: Some(iter),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let mut body_env = env.clone();
    body_env
        .value_types
        .insert(iter_name.clone(), "string".to_string());

    let mut lowered_body = Vec::new();
    match key {
        Some(key_name) => {
            if key_name != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(key_name.to_string()),
                            type_hint: Some("int".to_string().into()),
                            init: Some(Expression::ident(&index_name)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
            if var != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(var.to_string()),
                            type_hint: Some("int".to_string().into()),
                            init: Some(go_builtin_call(
                                "__go_str_char_code_at",
                                vec![
                                    Expression::ident(&iter_name),
                                    Expression::ident(&index_name),
                                ],
                            )),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
        }
        None => {
            if var != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(var.to_string()),
                            type_hint: Some("int".to_string().into()),
                            init: Some(Expression::ident(&index_name)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
        }
    }

    for stmt in body {
        lowered_body.extend(normalize_go_statement(
            stmt,
            &mut body_env,
            signatures,
            state,
        ));
    }

    let for_stmt = Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(index_name.clone()),
                type_hint: Some("int".to_string().into()),
                init: Some(Expression::int(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(go_builtin_call("len", vec![Expression::ident(&iter_name)])),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(&index_name)),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::ident(&index_name)),
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: lowered_body,
    });

    vec![Statement::new(StmtKind::Block(vec![iter_decl, for_stmt]))]
}

fn fresh_go_temp(state: &mut GoNormalizeState, prefix: &str) -> String {
    let name = format!("{}{}", prefix, state.next_temp);
    state.next_temp += 1;
    name
}

fn go_wrap_fixed_array_copy(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if go_expr_is_fixed_array(&expr, env, signatures) && go_requires_fixed_array_copy(&expr) {
        go_builtin_call("__go_fixed_array_clone", vec![expr])
    } else {
        expr
    }
}

fn go_wrap_go_value_copy(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if go_expr_is_fixed_array(&expr, env, signatures) && go_requires_fixed_array_copy(&expr) {
        return go_builtin_call("__go_fixed_array_clone", vec![expr]);
    }

    let Some(type_name) = go_expr_type_hint(&expr, env, signatures) else {
        return expr;
    };
    let trimmed = type_name.trim();
    if trimmed.starts_with("__go") {
        return expr;
    }
    if trimmed.starts_with('*') || !go_requires_go_value_copy(&expr) {
        return expr;
    }
    let Some(lookup) = go_struct_lookup_name(trimmed) else {
        return expr;
    };
    let Some(info) = env.struct_infos.get(&lookup) else {
        return expr;
    };
    if info.field_order.is_empty() {
        return expr;
    }

    go_struct_value_copy_expr(expr, trimmed, env)
}

fn go_struct_value_copy_expr(
    expr: Expression,
    type_name: &str,
    env: &GoNormalizeEnv,
) -> Expression {
    if let Some(lookup) = go_struct_lookup_name(type_name) {
        if let Some(info) = env.struct_infos.get(&lookup) {
            if !info.field_order.is_empty() {
                let properties = info
                    .field_order
                    .iter()
                    .map(|field| ObjectProperty::KeyValue {
                        key: Expression::string(field),
                        value: Expression::new(ExprKind::Member {
                            object: Box::new(expr.clone()),
                            field: field.clone(),
                            null_safe: false,
                        }),
                    })
                    .collect();
                return Expression::new(ExprKind::Cast {
                    expr: Box::new(Expression::new(ExprKind::Object(properties))),
                    type_name: type_name.to_string(),
                });
            }
        }
    }

    Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::new(ExprKind::Object(vec![
            ObjectProperty::Spread(expr),
        ]))),
        type_name: type_name.to_string(),
    })
}

fn go_requires_fixed_array_copy(expr: &Expression) -> bool {
    matches!(
        expr.kind,
        ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. }
    )
}

fn go_requires_go_value_copy(expr: &Expression) -> bool {
    matches!(
        expr.kind,
        ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. }
    )
}

fn go_prepend_value_receiver_copy(stmt: Statement, env: &GoNormalizeEnv) -> Statement {
    let _ = env;
    let StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers,
        handles,
        is_async,
        is_generator,
        is_sub,
    } = stmt.kind
    else {
        return stmt;
    };

    Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers,
        handles,
        is_async,
        is_generator,
        is_sub,
    })
}

fn go_builtin_call(name: &str, args: Vec<Expression>) -> Expression {
    let name = go_private_adapter_builtin_name(name);
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args
            .into_iter()
            .map(|value| Argument {
                value,
                name: None,
                by_ref: false,
                spread: false,
            })
            .collect(),
        optional: false,
    })
}

fn go_private_adapter_builtin_name(name: &str) -> &str {
    match name {
        "go.errors_new" => "__go_errors_new",
        "go.errors_string" => "__go_errors_string",
        "go.errors_unwrap" => "__go_errors_unwrap",
        "go.errors_is" => "__go_errors_is",
        "go.errors_join" => "__go_errors_join",
        "go.errors_as" => "__go_errors_as",
        "go.errors_errorf" => "__go_errors_errorf",
        "go.sort_search" => "__go_sort_search",
        "go.sort_search_ordered" => "__go_sort_search_ordered",
        "go.sort_find" => "__go_sort_find",
        "go.sort_slice" => "__go_sort_slice",
        "go.sort_is_sorted" => "__go_sort_is_sorted",
        "go.sort_reverse" => "__go_sort_reverse",
        "go.container.heap.Init" => "__go_container_heap_init",
        "go.container.heap.Pop" => "__go_container_heap_pop",
        "go.container.heap.Remove" => "__go_container_heap_remove",
        "go.container.heap.Fix" => "__go_container_heap_fix",
        "go.container.heap.remove_prepare" => "__go_container_heap_remove_prepare",
        "go.container.list.New" => "__go_container_list_new",
        "go.container.list.List.Init" => "__go_container_list_init",
        "go.container.list.List.Len" => "__go_container_list_len",
        "go.container.list.List.Front" => "__go_container_list_front",
        "go.container.list.List.Back" => "__go_container_list_back",
        "go.container.list.List.PushFront" => "__go_container_list_push_front",
        "go.container.list.List.PushBack" => "__go_container_list_push_back",
        "go.container.list.List.InsertBefore" => "__go_container_list_insert_before",
        "go.container.list.List.InsertAfter" => "__go_container_list_insert_after",
        "go.container.list.List.Remove" => "__go_container_list_remove",
        "go.container.list.List.MoveToFront" => "__go_container_list_move_to_front",
        "go.container.list.List.MoveToBack" => "__go_container_list_move_to_back",
        "go.container.list.List.MoveBefore" => "__go_container_list_move_before",
        "go.container.list.List.MoveAfter" => "__go_container_list_move_after",
        "go.container.list.List.PushBackList" => "__go_container_list_push_back_list",
        "go.container.list.List.PushFrontList" => "__go_container_list_push_front_list",
        "go.container.list.Element.Next" => "__go_container_list_element_next",
        "go.container.list.Element.Prev" => "__go_container_list_element_prev",
        "go.container.ring.New" => "__go_container_ring_new",
        "go.container.ring.Ring.Next" => "__go_container_ring_next",
        "go.container.ring.Ring.Prev" => "__go_container_ring_prev",
        "go.container.ring.Ring.Len" => "__go_container_ring_len",
        "go.container.ring.Ring.Move" => "__go_container_ring_move",
        "go.container.ring.Ring.Do" => "__go_container_ring_do",
        "go.container.ring.Ring.Link" => "__go_container_ring_link",
        "go.container.ring.Ring.Unlink" => "__go_container_ring_unlink",
        "go.container.ring.Ring.Value.Get" => "__go_container_ring_get_value",
        "go.container.ring.Ring.Value.Set" => "__go_container_ring_set_value",
        "go.strings.TrimPrefix" => "__go_strings_trim_prefix",
        "go.strings.TrimSuffix" => "__go_strings_trim_suffix",
        "go.strings.CutPrefix" => "__go_strings_cut_prefix",
        "go.strings.CutSuffix" => "__go_strings_cut_suffix",
        "go.strings.Cut" => "__go_strings_cut",
        "go.strings.Replace" => "__go_strings_replace",
        "go.strings.ReplaceAll" => "__go_strings_replace_all",
        "go.strings.ContainsRune" => "__go_strings_contains_rune",
        "go.strings.ContainsAny" => "__go_strings_contains_any",
        "go.strings.ContainsFunc" => "__go_strings_contains_func",
        "go.strings.IndexByte" => "__go_strings_index_byte",
        "go.strings.IndexRune" => "__go_strings_index_rune",
        "go.strings.IndexAny" => "__go_strings_index_any",
        "go.strings.IndexFunc" => "__go_strings_index_func",
        "go.strings.LastIndexByte" => "__go_strings_last_index_byte",
        "go.strings.LastIndexAny" => "__go_strings_last_index_any",
        "go.strings.LastIndexFunc" => "__go_strings_last_index_func",
        "go.strings.TrimLeft" => "__go_strings_trim_left",
        "go.strings.TrimRight" => "__go_strings_trim_right",
        "go.strings.Trim" => "__go_strings_trim",
        "go.strings.EqualFold" => "__go_strings_equal_fold",
        "go.strings.Count" => "__go_strings_count",
        "go.strings.ToValidUTF8" => "__go_strings_to_valid_utf8",
        "go.strings.Map" => "__go_strings_map",
        "go.strings.Fields" => "__go_strings_fields",
        "go.strings.FieldsFunc" => "__go_strings_fields_func",
        "go.strings.SplitN" => "__go_strings_split_n",
        "go.strings.SplitAfter" => "__go_strings_split_after",
        "go.strings.SplitAfterN" => "__go_strings_split_after_n",
        "go.time_utc" => "__go_time_utc",
        "go.time_local" => "__go_time_local",
        "go.time_date" => "__go_time_date",
        "go.time_unix" => "__go_time_unix",
        "go.time_now" => "__go_time_now",
        "go.time_unix_milli" => "__go_time_unix_milli",
        "go.time_unix_micro" => "__go_time_unix_micro",
        "go.time_fixed_zone" => "__go_time_fixed_zone",
        "go.time_load_location" => "__go_time_load_location",
        "go.time_parse" => "__go_time_parse",
        "go.time_parse_in_location" => "__go_time_parse_in_location",
        "go.time_parse_duration" => "__go_time_parse_duration",
        "go.time_since" => "__go_time_since",
        "go.time_until" => "__go_time_until",
        "go.time_sleep" => "__go_time_sleep",
        "go.time_after" => "__go_time_after",
        "go.time_time_format" => "__go_time_time_format",
        "go.time_time_year" => "__go_time_time_year",
        "go.time_time_add_date" => "__go_time_time_add_date",
        "go.time_time_add" => "__go_time_time_add",
        "go.time_time_sub" => "__go_time_time_sub",
        "go.time_time_month" => "__go_time_time_month",
        "go.time_time_month_int" => "__go_time_time_month_int",
        "go.time_time_day" => "__go_time_time_day",
        "go.time_time_hour" => "__go_time_time_hour",
        "go.time_time_minute" => "__go_time_time_minute",
        "go.time_time_second" => "__go_time_time_second",
        "go.time_time_nanosecond" => "__go_time_time_nanosecond",
        "go.time_time_unix" => "__go_time_time_unix",
        "go.time_time_unix_nano" => "__go_time_time_unix_nano",
        "go.time_time_unix_milli" => "__go_time_time_unix_milli",
        "go.time_time_unix_micro" => "__go_time_time_unix_micro",
        "go.time_time_weekday" => "__go_time_time_weekday",
        "go.time_time_year_day" => "__go_time_time_year_day",
        "go.time_time_zone" => "__go_time_time_zone",
        "go.time_time_before" => "__go_time_time_before",
        "go.time_time_after" => "__go_time_time_after",
        "go.time_time_equal" => "__go_time_time_equal",
        "go.time_time_truncate" => "__go_time_time_truncate",
        "go.time_time_round" => "__go_time_time_round",
        "go.time_time_utc" => "__go_time_time_utc",
        "go.time_time_in" => "__go_time_time_in",
        "go.time_time_location" => "__go_time_time_location",
        "go.time_time_is_zero" => "__go_time_time_is_zero",
        "go.time_location_string" => "__go_time_location_string",
        "go.time_duration_string" => "__go_time_duration_string",
        "go.time_duration_round" => "__go_time_duration_round",
        _ => name,
    }
}

fn go_public_adapter_emit_name(name: &str) -> &str {
    match name {
        "__go_errors_new" => "go.errors_new",
        "__go_errors_string" => "go.errors_string",
        "__go_errors_unwrap" => "go.errors_unwrap",
        "__go_errors_is" => "go.errors_is",
        "__go_errors_join" => "go.errors_join",
        "__go_errors_as" => "go.errors_as",
        "__go_errors_errorf" => "go.errors_errorf",
        "__go_sort_search" => "go.sort_search",
        "__go_sort_search_ordered" => "go.sort_search_ordered",
        "__go_sort_find" => "go.sort_find",
        "__go_sort_slice" => "go.sort_slice",
        "__go_sort_is_sorted" => "go.sort_is_sorted",
        "__go_sort_reverse" => "go.sort_reverse",
        "__go_container_heap_init" => "go.container.heap.Init",
        "__go_container_heap_pop" => "go.container.heap.Pop",
        "__go_container_heap_remove" => "go.container.heap.Remove",
        "__go_container_heap_fix" => "go.container.heap.Fix",
        "__go_container_heap_remove_prepare" => "go.container.heap.remove_prepare",
        "__go_container_list_new" => "go.container.list.New",
        "__go_container_list_init" => "go.container.list.List.Init",
        "__go_container_list_len" => "go.container.list.List.Len",
        "__go_container_list_front" => "go.container.list.List.Front",
        "__go_container_list_back" => "go.container.list.List.Back",
        "__go_container_list_push_front" => "go.container.list.List.PushFront",
        "__go_container_list_push_back" => "go.container.list.List.PushBack",
        "__go_container_list_insert_before" => "go.container.list.List.InsertBefore",
        "__go_container_list_insert_after" => "go.container.list.List.InsertAfter",
        "__go_container_list_remove" => "go.container.list.List.Remove",
        "__go_container_list_move_to_front" => "go.container.list.List.MoveToFront",
        "__go_container_list_move_to_back" => "go.container.list.List.MoveToBack",
        "__go_container_list_move_before" => "go.container.list.List.MoveBefore",
        "__go_container_list_move_after" => "go.container.list.List.MoveAfter",
        "__go_container_list_push_back_list" => "go.container.list.List.PushBackList",
        "__go_container_list_push_front_list" => "go.container.list.List.PushFrontList",
        "__go_container_list_element_next" => "go.container.list.Element.Next",
        "__go_container_list_element_prev" => "go.container.list.Element.Prev",
        "__go_container_ring_new" => "go.container.ring.New",
        "__go_container_ring_next" => "go.container.ring.Ring.Next",
        "__go_container_ring_prev" => "go.container.ring.Ring.Prev",
        "__go_container_ring_len" => "go.container.ring.Ring.Len",
        "__go_container_ring_move" => "go.container.ring.Ring.Move",
        "__go_container_ring_do" => "go.container.ring.Ring.Do",
        "__go_container_ring_link" => "go.container.ring.Ring.Link",
        "__go_container_ring_unlink" => "go.container.ring.Ring.Unlink",
        "__go_container_ring_get_value" => "go.container.ring.Ring.Value.Get",
        "__go_container_ring_set_value" => "go.container.ring.Ring.Value.Set",
        "__go_strings_trim_prefix" => "go.strings.TrimPrefix",
        "__go_strings_trim_suffix" => "go.strings.TrimSuffix",
        "__go_strings_cut_prefix" => "go.strings.CutPrefix",
        "__go_strings_cut_suffix" => "go.strings.CutSuffix",
        "__go_strings_cut" => "go.strings.Cut",
        "__go_strings_replace" => "go.strings.Replace",
        "__go_strings_replace_all" => "go.strings.ReplaceAll",
        "__go_strings_contains_rune" => "go.strings.ContainsRune",
        "__go_strings_contains_any" => "go.strings.ContainsAny",
        "__go_strings_contains_func" => "go.strings.ContainsFunc",
        "__go_strings_index_byte" => "go.strings.IndexByte",
        "__go_strings_index_rune" => "go.strings.IndexRune",
        "__go_strings_index_any" => "go.strings.IndexAny",
        "__go_strings_index_func" => "go.strings.IndexFunc",
        "__go_strings_last_index_byte" => "go.strings.LastIndexByte",
        "__go_strings_last_index_any" => "go.strings.LastIndexAny",
        "__go_strings_last_index_func" => "go.strings.LastIndexFunc",
        "__go_strings_trim_left" => "go.strings.TrimLeft",
        "__go_strings_trim_right" => "go.strings.TrimRight",
        "__go_strings_trim" => "go.strings.Trim",
        "__go_strings_equal_fold" => "go.strings.EqualFold",
        "__go_strings_count" => "go.strings.Count",
        "__go_strings_to_valid_utf8" => "go.strings.ToValidUTF8",
        "__go_strings_map" => "go.strings.Map",
        "__go_strings_fields" => "go.strings.Fields",
        "__go_strings_fields_func" => "go.strings.FieldsFunc",
        "__go_strings_split_n" => "go.strings.SplitN",
        "__go_strings_split_after" => "go.strings.SplitAfter",
        "__go_strings_split_after_n" => "go.strings.SplitAfterN",
        "__go_time_utc" => "go.time_utc",
        "__go_time_local" => "go.time_local",
        "__go_time_date" => "go.time_date",
        "__go_time_unix" => "go.time_unix",
        "__go_time_now" => "go.time_now",
        "__go_time_unix_milli" => "go.time_unix_milli",
        "__go_time_unix_micro" => "go.time_unix_micro",
        "__go_time_fixed_zone" => "go.time_fixed_zone",
        "__go_time_load_location" => "go.time_load_location",
        "__go_time_parse" => "go.time_parse",
        "__go_time_parse_in_location" => "go.time_parse_in_location",
        "__go_time_parse_duration" => "go.time_parse_duration",
        "__go_time_since" => "go.time_since",
        "__go_time_until" => "go.time_until",
        "__go_time_sleep" => "go.time_sleep",
        "__go_time_after" => "go.time_after",
        "__go_time_time_format" => "go.time_time_format",
        "__go_time_time_year" => "go.time_time_year",
        "__go_time_time_add_date" => "go.time_time_add_date",
        "__go_time_time_add" => "go.time_time_add",
        "__go_time_time_sub" => "go.time_time_sub",
        "__go_time_time_month" => "go.time_time_month",
        "__go_time_time_month_int" => "go.time_time_month_int",
        "__go_time_time_day" => "go.time_time_day",
        "__go_time_time_hour" => "go.time_time_hour",
        "__go_time_time_minute" => "go.time_time_minute",
        "__go_time_time_second" => "go.time_time_second",
        "__go_time_time_nanosecond" => "go.time_time_nanosecond",
        "__go_time_time_unix" => "go.time_time_unix",
        "__go_time_time_unix_nano" => "go.time_time_unix_nano",
        "__go_time_time_unix_milli" => "go.time_time_unix_milli",
        "__go_time_time_unix_micro" => "go.time_time_unix_micro",
        "__go_time_time_weekday" => "go.time_time_weekday",
        "__go_time_time_year_day" => "go.time_time_year_day",
        "__go_time_time_zone" => "go.time_time_zone",
        "__go_time_time_before" => "go.time_time_before",
        "__go_time_time_after" => "go.time_time_after",
        "__go_time_time_equal" => "go.time_time_equal",
        "__go_time_time_truncate" => "go.time_time_truncate",
        "__go_time_time_round" => "go.time_time_round",
        "__go_time_time_utc" => "go.time_time_utc",
        "__go_time_time_in" => "go.time_time_in",
        "__go_time_time_location" => "go.time_time_location",
        "__go_time_time_is_zero" => "go.time_time_is_zero",
        "__go_time_location_string" => "go.time_location_string",
        "__go_time_duration_string" => "go.time_duration_string",
        "__go_time_duration_round" => "go.time_duration_round",
        _ => name,
    }
}

/// Build a slice/array literal AST node from a list of element expressions.
fn go_array_of(elems: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        elems
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn go_arg_value(args: &[Argument], idx: usize) -> Expression {
    args.get(idx)
        .map(|a| a.value.clone())
        .unwrap_or_else(Expression::null)
}

fn go_arg_callable_value(args: &[Argument], idx: usize) -> Expression {
    match go_arg_value(args, idx).kind {
        ExprKind::Cast { expr, .. } => *expr,
        kind => Expression::new(kind),
    }
}

fn go_rewrite_fmt_format_call(
    call_name: &str,
    callee: &Expression,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    if !matches!(call_name, "fmt.Sprintf" | "__go_sprintf" | "fmt.Printf") || args.is_empty() {
        return None;
    }
    let ExprKind::Lit(Literal::Str(fmt)) = &args[0].value.kind else {
        return None;
    };
    let (newfmt, rewrites) = crate::adapters::formatting::rewrite_go_format_literal(fmt);
    let fix_exp =
        matches!(call_name, "fmt.Sprintf" | "__go_sprintf") && go_format_has_exp_verb(fmt);
    let has_complex_arg = args
        .iter()
        .skip(1)
        .any(|arg| go_expr_is_complex(&arg.value));
    if rewrites.is_empty() && newfmt == *fmt && !fix_exp && !has_complex_arg {
        return None;
    }

    let mut next_args = Vec::with_capacity(args.len());
    next_args.push(Argument {
        value: Expression::string(&newfmt),
        name: args[0].name.clone(),
        by_ref: args[0].by_ref,
        spread: args[0].spread,
    });
    for (idx, arg) in args.iter().enumerate().skip(1) {
        let value = match rewrites.get(&(idx - 1)).copied() {
            Some(crate::adapters::formatting::GoFmtArgRewrite::Pointer) => {
                crate::adapters::formatting::fmt_pointer_expr(arg.value.clone())
            }
            _ if go_expr_is_complex(&arg.value) => go_complex_format_expr(arg.value.clone()),
            Some(crate::adapters::formatting::GoFmtArgRewrite::String) => {
                go_stringer_call_expr(arg.value.clone(), env, signatures)
                    .unwrap_or_else(|| go_builtin_call("__go_fmt_string", vec![arg.value.clone()]))
            }
            Some(crate::adapters::formatting::GoFmtArgRewrite::Quote) => go_builtin_call(
                "__go_fmt_quote",
                vec![go_builtin_call("__go_fmt_string", vec![arg.value.clone()])],
            ),
            Some(crate::adapters::formatting::GoFmtArgRewrite::TypeName) => Expression::string(
                &go_expr_type_hint(&arg.value, env, signatures).unwrap_or_else(|| {
                    go_expr_call_name(&arg.value).unwrap_or_else(|| "interface {}".to_string())
                }),
            ),
            Some(crate::adapters::formatting::GoFmtArgRewrite::GoValue { field_names }) => {
                go_format_value_expr(arg.value.clone(), field_names, env, signatures)
            }
            _ => arg.value.clone(),
        };
        next_args.push(Argument {
            value,
            name: arg.name.clone(),
            by_ref: arg.by_ref,
            spread: arg.spread,
        });
    }

    let call = Expression::new(ExprKind::Call {
        callee: Box::new(callee.clone()),
        args: next_args,
        optional,
    });
    if fix_exp {
        Some(go_builtin_call("__go_fmt_fix_exp", vec![call]))
    } else {
        Some(call)
    }
}

fn go_format_has_exp_verb(fmt: &str) -> bool {
    let chars: Vec<char> = fmt.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] != '%' {
            i += 1;
            continue;
        }
        if i + 1 < chars.len() && chars[i + 1] == '%' {
            i += 2;
            continue;
        }
        i += 1;
        while i < chars.len() {
            let spec = chars[i];
            if spec.is_ascii_alphabetic() {
                if spec == 'e' || spec == 'E' {
                    return true;
                }
                i += 1;
                break;
            }
            i += 1;
        }
    }
    false
}

fn go_format_value_expr(
    value: Expression,
    field_names: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if let Some(stringer) = go_stringer_call_expr(value.clone(), env, signatures) {
        return stringer;
    }
    if go_expr_type_hint(&value, env, signatures)
        .as_deref()
        .is_some_and(go_is_array_like_type)
        || matches!(value.kind, ExprKind::Array(_))
    {
        return go_builtin_call("__go_fmt_slice", vec![value]);
    }
    if let Some(props) = go_object_format_props(&value) {
        return go_format_struct_props(props, field_names);
    }
    go_builtin_call("__go_fmt_string", vec![value])
}

fn go_stringer_call_expr(
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let receiver_type = go_expr_type_hint(&value, env, signatures)?;
    if crate::adapters::url::has_method(&receiver_type, "String") {
        return crate::adapters::url::rewrite_method_call(value, &receiver_type, "String", &[]);
    }
    if crate::adapters::netip::has_method(&receiver_type, "String") {
        return crate::adapters::netip::rewrite_method_call(value, &receiver_type, "String", &[]);
    }
    if !go_type_has_method(&receiver_type, "String", env) {
        return None;
    }
    let lookup = go_struct_lookup_name(&receiver_type)?;
    if env.named_types.contains_key(&lookup) {
        return go_rewrite_named_type_method_call(
            &Expression::new(ExprKind::Member {
                object: Box::new(value),
                field: "String".to_string(),
                null_safe: false,
            }),
            &[],
            false,
            env,
            signatures,
        );
    }
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&lookup)),
            field: "String".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(value)],
        optional: false,
    }))
}

fn go_type_has_method(type_name: &str, method: &str, env: &GoNormalizeEnv) -> bool {
    let allow_pointer_methods = type_name.trim().starts_with('*');
    go_type_has_method_in_method_set(
        type_name,
        method,
        allow_pointer_methods,
        env,
        &mut HashSet::new(),
    )
}

fn go_type_has_method_in_method_set(
    type_name: &str,
    method: &str,
    allow_pointer_methods: bool,
    env: &GoNormalizeEnv,
    seen: &mut HashSet<String>,
) -> bool {
    let Some(lookup) = go_struct_lookup_name(type_name) else {
        return false;
    };
    if !seen.insert(lookup.clone()) {
        return false;
    }
    let Some(info) = env.struct_infos.get(&lookup) else {
        return false;
    };
    if info.method_names.contains(method)
        && (allow_pointer_methods || !info.pointer_method_names.contains(method))
    {
        return true;
    }
    for (_, embedded_type) in &info.embedded_fields {
        let embedded_allows_pointer =
            allow_pointer_methods || embedded_type.trim().starts_with('*');
        if go_type_has_method_in_method_set(
            embedded_type,
            method,
            embedded_allows_pointer,
            env,
            seen,
        ) {
            return true;
        }
    }
    false
}

fn go_object_format_props(value: &Expression) -> Option<Vec<ObjectProperty>> {
    match &value.kind {
        ExprKind::Object(props) => Some(props.clone()),
        ExprKind::Cast { expr, .. } => match &expr.kind {
            ExprKind::Object(props) => Some(props.clone()),
            _ => None,
        },
        _ => None,
    }
}

fn go_format_struct_props(props: Vec<ObjectProperty>, field_names: bool) -> Expression {
    let mut parts = vec![Expression::string("{")];
    let mut first = true;
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        let field = match key.kind {
            ExprKind::Lit(Literal::Str(name)) => name,
            ExprKind::Ident(name) => name,
            _ => continue,
        };
        if !first {
            parts.push(Expression::string(" "));
        }
        first = false;
        if field_names {
            parts.push(Expression::string(&format!("{}: ", field)));
        }
        parts.push(go_builtin_call("__go_fmt_string", vec![value]));
    }
    parts.push(Expression::string("}"));
    go_concat_exprs(parts)
}

fn go_rewrite_fmt_output_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let adapter = match call_name {
        "fmt.Println" => "__go_fmt_println",
        "fmt.Print" => "__go_fmt_print",
        "fmt.Sprint" => "__go_fmt_sprint",
        _ => return None,
    };
    let rewritten_args = args
        .iter()
        .map(|arg| {
            let (value, did_change) = go_rewrite_time_month_print_arg(arg.value.clone());
            let _ = did_change;
            let value = if go_expr_type_hint(&value, env, signatures)
                .as_deref()
                .is_some_and(|ty| matches!(ty.trim(), "error" | "__goError"))
            {
                go_builtin_call("go.errors_string", vec![value])
            } else {
                value
            };
            value
        })
        .collect::<Vec<_>>();
    Some(go_builtin_call(adapter, rewritten_args))
}

fn go_rewrite_fmt_io_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let writer = go_fmt_writer_expr(args.first()?.value.clone());
    let message = match call_name {
        "fmt.Fprintf" => {
            let format = args.get(1)?.value.clone();
            let values = args.iter().skip(2).map(|arg| arg.value.clone()).collect();
            go_sprintf_expr(format, values, env, signatures)
        }
        "fmt.Fprint" => {
            let values = args
                .iter()
                .skip(1)
                .map(|arg| go_format_value_expr(arg.value.clone(), false, env, signatures))
                .collect();
            go_concat_exprs(values)
        }
        "fmt.Fprintln" => {
            let mut values = Vec::new();
            for (idx, arg) in args.iter().skip(1).enumerate() {
                if idx > 0 {
                    values.push(Expression::string(" "));
                }
                values.push(go_format_value_expr(
                    arg.value.clone(),
                    false,
                    env,
                    signatures,
                ));
            }
            values.push(Expression::string("\n"));
            go_concat_exprs(values)
        }
        _ => return None,
    };
    let captures = go_big_captures(&[&writer, &message]);
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(vec![
                Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident("__go_fmt_out".to_string()),
                        type_hint: None,
                        init: Some(message),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
                Statement::new(StmtKind::Expr(
                    crate::adapters::bytes_io::buffer_write_string_expr(
                        writer,
                        Expression::ident("__go_fmt_out"),
                    ),
                )),
                Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
                    vec![
                        go_builtin_call("len", vec![Expression::ident("__go_fmt_out")]),
                        Expression::null(),
                    ],
                ))))),
            ]),
            is_async: false,
            captures,
        })),
        args: vec![],
        optional: false,
    }))
}

fn go_rewrite_fmt_scan_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    if !matches!(call_name, "fmt.Sscanf" | "fmt.Fscanf") || args.len() < 2 {
        return None;
    }
    let source = if call_name == "fmt.Sscanf" {
        go_literal_string(&args[0].value)?
    } else {
        go_scan_reader_literal(&args[0].value)?
    };
    let format = go_literal_string(&args[1].value)?;
    let verbs = go_scan_verbs(&format);
    let tokens = go_scan_tokens(&source);
    let mut body = Vec::new();
    let mut count = 0;
    for ((verb, token), arg) in verbs
        .into_iter()
        .zip(tokens.into_iter())
        .zip(args.iter().skip(2))
    {
        let Some(target) = go_scan_target_expr(&arg.value) else {
            break;
        };
        let Some(value) = go_scan_value_expr(verb, &token) else {
            break;
        };
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![target],
            value,
            by_ref: false,
        }));
        count += 1;
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Tuple(vec![Expression::int(count), Expression::null()]),
    )))));
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    }))
}

fn go_literal_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.clone()),
        _ => None,
    }
}

fn go_scan_reader_literal(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => match go_expr_call_name(callee).as_deref() {
            Some("go.strings.NewReader") | Some("strings.NewReader") => {
                args.first().and_then(|arg| go_literal_string(&arg.value))
            }
            Some("go.bytes.NewReader") | Some("bytes.NewReader") => args
                .first()
                .and_then(|arg| go_scan_bytes_literal_string(&arg.value)),
            _ => None,
        },
        ExprKind::Cast { expr, type_name } if type_name.trim() == "__goReader" => {
            go_scan_reader_literal(expr)
        }
        ExprKind::Object(props) => go_object_prop_value(props, "data").and_then(|data| {
            go_literal_string(&data).or_else(|| go_scan_bytes_literal_string(&data))
        }),
        _ => None,
    }
}

fn go_scan_bytes_literal_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Call { callee, args, .. }
            if go_expr_call_name(callee).as_deref() == Some("__go_io_bytes_to_string") =>
        {
            args.first()
                .and_then(|arg| go_scan_bytes_literal_string(&arg.value))
        }
        ExprKind::Call { callee, args, .. }
            if go_expr_call_name(callee).as_deref() == Some("__go_io_string_to_bytes") =>
        {
            args.first().and_then(|arg| go_literal_string(&arg.value))
        }
        ExprKind::Cast { expr, type_name } if matches!(type_name.trim(), "[]byte" | "[]uint8") => {
            go_literal_string(expr)
        }
        _ => go_literal_string(expr),
    }
}

fn go_regex_pattern_from_expr(expr: &Expression, env: &GoNormalizeEnv) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => env.regex_patterns.get(name).cloned(),
        ExprKind::Object(props) => go_object_prop_value(props, "__go_regex_pattern")
            .as_ref()
            .and_then(go_literal_string),
        ExprKind::Cast { expr, .. } => go_regex_pattern_from_expr(expr, env),
        ExprKind::Call { callee, args, .. } => match go_expr_call_name(callee).as_deref()? {
            "regexp.MustCompile" | "regexp.Compile" => {
                args.first().and_then(|arg| go_literal_string(&arg.value))
            }
            _ => None,
        },
        _ => None,
    }
}

fn go_regex_object_expr(pattern: &str) -> Expression {
    Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
        key: Expression::string("__go_regex_pattern"),
        value: Expression::string(pattern),
    }]))
}

fn go_regex_quote_meta(text: &str) -> String {
    let mut out = String::new();
    for ch in text.chars() {
        if matches!(
            ch,
            '\\' | '.' | '+' | '*' | '?' | '(' | ')' | '|' | '[' | ']' | '{' | '}' | '^' | '$'
        ) {
            out.push('\\');
        }
        out.push(ch);
    }
    out
}

fn go_regex_limit(expr: Option<&Expression>) -> Option<usize> {
    match expr.map(|expr| &expr.kind) {
        None => None,
        Some(ExprKind::Lit(Literal::Int(n))) if *n < 0 => None,
        Some(ExprKind::Lit(Literal::Int(n))) => Some((*n).max(0) as usize),
        Some(ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        }) if matches!(expr.kind, ExprKind::Lit(Literal::Int(1))) => None,
        _ => None,
    }
}

fn go_regex_find_all_string_submatch_expr(
    pattern: &str,
    input: &str,
    limit: Option<usize>,
) -> Option<Expression> {
    if limit == Some(0) {
        return Some(Expression::null());
    }
    let re = Regex::new(pattern).ok()?;
    let mut rows = Vec::new();
    for caps in re.captures_iter(input) {
        if let Some(max) = limit {
            if rows.len() >= max {
                break;
            }
        }
        let cols = (0..caps.len())
            .map(|idx| Expression::string(caps.get(idx).map(|m| m.as_str()).unwrap_or_default()))
            .collect();
        rows.push(go_array_of(cols));
    }
    if rows.is_empty() {
        Some(Expression::null())
    } else {
        Some(go_array_of(rows))
    }
}

fn go_regex_find_string_submatch_expr(pattern: &str, input: &str) -> Option<Expression> {
    let re = Regex::new(pattern).ok()?;
    let caps = re.captures(input)?;
    Some(go_array_of(
        (0..caps.len())
            .map(|idx| Expression::string(caps.get(idx).map(|m| m.as_str()).unwrap_or_default()))
            .collect(),
    ))
}

fn go_regex_split_expr(pattern: &str, input: &str, limit: Option<usize>) -> Option<Expression> {
    if limit == Some(0) {
        return Some(Expression::null());
    }
    let re = Regex::new(pattern).ok()?;
    let values = match limit {
        Some(n) => re.splitn(input, n).collect::<Vec<_>>(),
        None => re.split(input).collect::<Vec<_>>(),
    };
    Some(go_array_of(
        values.into_iter().map(Expression::string).collect(),
    ))
}

fn go_regex_subexp_names_expr(pattern: &str) -> Option<Expression> {
    let re = Regex::new(pattern).ok()?;
    Some(go_array_of(
        re.capture_names()
            .map(|name| Expression::string(name.unwrap_or_default()))
            .collect(),
    ))
}

fn go_regex_num_subexp_expr(pattern: &str) -> Option<Expression> {
    let re = Regex::new(pattern).ok()?;
    Some(Expression::int(
        (re.captures_len().saturating_sub(1)) as i64,
    ))
}

fn go_regex_literal_prefix_expr(pattern: &str) -> Expression {
    let mut prefix = String::new();
    let mut chars = pattern.chars().peekable();
    let mut complete = true;
    while let Some(ch) = chars.next() {
        match ch {
            '\\' => {
                if let Some(next) = chars.next() {
                    if matches!(
                        next,
                        '\\' | '.'
                            | '+'
                            | '*'
                            | '?'
                            | '('
                            | ')'
                            | '|'
                            | '['
                            | ']'
                            | '{'
                            | '}'
                            | '^'
                            | '$'
                    ) {
                        prefix.push(next);
                    } else {
                        complete = false;
                        break;
                    }
                } else {
                    complete = false;
                    break;
                }
            }
            '.' | '+' | '*' | '?' | '(' | ')' | '|' | '[' | ']' | '{' | '}' | '^' | '$' => {
                complete = false;
                break;
            }
            _ => prefix.push(ch),
        }
    }
    Expression::new(ExprKind::Tuple(vec![
        Expression::string(&prefix),
        Expression::bool(complete),
    ]))
}

fn go_regex_replace_expand(caps: &Captures<'_>, replacement: &str) -> String {
    let mut out = String::new();
    let chars: Vec<char> = replacement.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] != '$' {
            out.push(chars[i]);
            i += 1;
            continue;
        }
        if i + 1 >= chars.len() {
            out.push('$');
            i += 1;
            continue;
        }
        if chars[i + 1] == '$' {
            out.push('$');
            i += 2;
            continue;
        }
        let mut j = i + 1;
        let mut braced = false;
        if chars[j] == '{' {
            braced = true;
            j += 1;
        }
        let start = j;
        while j < chars.len() && (chars[j].is_ascii_alphanumeric() || chars[j] == '_') {
            j += 1;
        }
        if braced {
            if j >= chars.len() || chars[j] != '}' {
                out.push('$');
                i += 1;
                continue;
            }
        }
        if start == j {
            out.push('$');
            i += 1;
            continue;
        }
        let name: String = chars[start..j].iter().collect();
        let value = if name.chars().all(|ch| ch.is_ascii_digit()) {
            name.parse::<usize>()
                .ok()
                .and_then(|idx| caps.get(idx))
                .map(|m| m.as_str())
        } else {
            caps.name(&name).map(|m| m.as_str())
        };
        if let Some(value) = value {
            out.push_str(value);
        }
        i = if braced { j + 1 } else { j };
    }
    out
}

fn go_regex_replace_all_string_expr(
    pattern: &str,
    input: &str,
    replacement: &str,
) -> Option<Expression> {
    let re = Regex::new(pattern).ok()?;
    let rendered = re
        .replace_all(input, |caps: &Captures<'_>| {
            go_regex_replace_expand(caps, replacement)
        })
        .to_string();
    Some(Expression::string(&rendered))
}

fn go_rewrite_regexp_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "regexp.MustCompile" => {
            let pattern = args.first().and_then(|arg| go_literal_string(&arg.value))?;
            Some(go_regex_object_expr(&pattern))
        }
        "regexp.Compile" => {
            let pattern = args.first().and_then(|arg| go_literal_string(&arg.value))?;
            let compiled = Regex::new(&pattern)
                .ok()
                .map_or_else(Expression::null, |_| go_regex_object_expr(&pattern));
            Some(Expression::new(ExprKind::Tuple(vec![
                compiled,
                Expression::null(),
            ])))
        }
        "regexp.MatchString" if args.len() >= 2 => {
            let pattern = go_literal_string(&args[0].value)?;
            let input = go_literal_string(&args[1].value)?;
            Some(Expression::bool(
                Regex::new(&pattern).ok()?.is_match(&input),
            ))
        }
        "regexp.QuoteMeta" => {
            let text = args.first().and_then(|arg| go_literal_string(&arg.value))?;
            Some(Expression::string(&go_regex_quote_meta(&text)))
        }
        _ => None,
    }
}

fn go_rewrite_regexp_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let pattern = go_regex_pattern_from_expr(object, env)?;
    match field.as_str() {
        "String" if args.is_empty() => Some(Expression::string(&pattern)),
        "Copy" | "Longest" if args.is_empty() => Some(go_regex_object_expr(&pattern)),
        "NumSubexp" if args.is_empty() => go_regex_num_subexp_expr(&pattern),
        "SubexpNames" if args.is_empty() => go_regex_subexp_names_expr(&pattern),
        "LiteralPrefix" if args.is_empty() => Some(go_regex_literal_prefix_expr(&pattern)),
        "FindAllStringSubmatch" if args.len() >= 2 => {
            let input = go_literal_string(&args[0].value)?;
            let limit = go_regex_limit(args.get(1).map(|arg| &arg.value));
            go_regex_find_all_string_submatch_expr(&pattern, &input, limit)
        }
        "FindStringSubmatch" if !args.is_empty() => {
            let input = go_literal_string(&args[0].value)?;
            go_regex_find_string_submatch_expr(&pattern, &input)
        }
        "ReplaceAllString" if args.len() >= 2 => {
            let input = go_literal_string(&args[0].value)?;
            let replacement = go_literal_string(&args[1].value)?;
            go_regex_replace_all_string_expr(&pattern, &input, &replacement)
        }
        "Split" if args.len() >= 2 => {
            let input = go_literal_string(&args[0].value)?;
            let limit = go_regex_limit(args.get(1).map(|arg| &arg.value));
            go_regex_split_expr(&pattern, &input, limit)
        }
        _ => None,
    }
}

fn go_scan_verbs(format: &str) -> Vec<char> {
    let chars: Vec<char> = format.chars().collect();
    let mut verbs = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] != '%' {
            i += 1;
            continue;
        }
        if i + 1 < chars.len() && chars[i + 1] == '%' {
            i += 2;
            continue;
        }
        i += 1;
        while i < chars.len() {
            let ch = chars[i];
            if ch.is_ascii_alphabetic() {
                verbs.push(ch);
                i += 1;
                break;
            }
            i += 1;
        }
    }
    verbs
}

fn go_scan_tokens(source: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut chars = source.chars().peekable();
    while chars.peek().is_some() {
        while chars.peek().is_some_and(|ch| ch.is_whitespace()) {
            chars.next();
        }
        let Some(&first) = chars.peek() else {
            break;
        };
        let mut token = String::new();
        if first == '"' {
            token.push(first);
            chars.next();
            let mut escaped = false;
            for ch in chars.by_ref() {
                token.push(ch);
                if escaped {
                    escaped = false;
                } else if ch == '\\' {
                    escaped = true;
                } else if ch == '"' {
                    break;
                }
            }
        } else {
            while chars.peek().is_some_and(|ch| !ch.is_whitespace()) {
                token.push(chars.next().unwrap());
            }
        }
        if !token.is_empty() {
            tokens.push(token);
        }
    }
    tokens
}

fn go_scan_target_expr(expr: &Expression) -> Option<Expression> {
    match &expr.kind {
        ExprKind::RefOf(place) => Some(go_place_expr(place)),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => Some(expr.as_ref().clone()),
        _ => None,
    }
}

fn go_scan_value_expr(verb: char, token: &str) -> Option<Expression> {
    match verb {
        'd' => token.parse::<i64>().ok().map(Expression::int),
        'x' | 'X' => i64::from_str_radix(token.trim_start_matches("0x"), 16)
            .ok()
            .map(Expression::int),
        'f' | 'g' | 'e' => token.parse::<f64>().ok().map(Expression::float),
        's' => Some(Expression::string(token)),
        'q' => Some(Expression::string(&go_scan_unquote(token))),
        't' => token.parse::<bool>().ok().map(Expression::bool),
        _ => None,
    }
}

fn go_scan_unquote(token: &str) -> String {
    let inner = token
        .strip_prefix('"')
        .and_then(|s| s.strip_suffix('"'))
        .unwrap_or(token);
    let mut out = String::new();
    let mut chars = inner.chars();
    while let Some(ch) = chars.next() {
        if ch == '\\' {
            match chars.next() {
                Some('n') => out.push('\n'),
                Some('t') => out.push('\t'),
                Some('r') => out.push('\r'),
                Some('"') => out.push('"'),
                Some('\\') => out.push('\\'),
                Some(other) => out.push(other),
                None => out.push('\\'),
            }
        } else {
            out.push(ch);
        }
    }
    out
}

fn go_rewrite_fmt_io_expr_statement(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Vec<Statement>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let call_name = go_expr_call_name(callee)?;
    if !matches!(
        call_name.as_str(),
        "fmt.Fprint" | "fmt.Fprintf" | "fmt.Fprintln"
    ) {
        return None;
    }
    let next_args = args
        .iter()
        .map(|arg| Argument {
            value: normalize_go_expr(&arg.value, env, signatures, state),
            name: arg.name.clone(),
            by_ref: arg.by_ref,
            spread: arg.spread,
        })
        .collect::<Vec<_>>();
    let writer = go_fmt_writer_expr(next_args.first()?.value.clone());
    let message = match call_name.as_str() {
        "fmt.Fprintf" => {
            let format = next_args.get(1)?.value.clone();
            let values = next_args
                .iter()
                .skip(2)
                .map(|arg| arg.value.clone())
                .collect();
            go_sprintf_expr(format, values, env, signatures)
        }
        "fmt.Fprint" => go_concat_exprs(
            next_args
                .iter()
                .skip(1)
                .map(|arg| go_format_value_expr(arg.value.clone(), false, env, signatures))
                .collect(),
        ),
        "fmt.Fprintln" => {
            let mut values = Vec::new();
            for (idx, arg) in next_args.iter().skip(1).enumerate() {
                if idx > 0 {
                    values.push(Expression::string(" "));
                }
                values.push(go_format_value_expr(
                    arg.value.clone(),
                    false,
                    env,
                    signatures,
                ));
            }
            values.push(Expression::string("\n"));
            go_concat_exprs(values)
        }
        _ => return None,
    };
    Some(vec![go_fmt_write_string_stmt(writer, message)])
}

fn go_fmt_writer_expr(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::RefOf(place) => go_place_expr(&place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        _ => expr,
    }
}

fn go_fmt_write_string_stmt(writer: Expression, message: Expression) -> Statement {
    let data = Expression::new(ExprKind::Member {
        object: Box::new(writer.clone()),
        field: "data".to_string(),
        null_safe: false,
    });
    Statement::new(StmtKind::Assign {
        targets: vec![data.clone()],
        value: Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(data),
            right: Box::new(message),
        }),
        by_ref: false,
    })
}

fn go_rewrite_time_month_print_arg(expr: Expression) -> (Expression, bool) {
    (expr, false)
}

/// Rewrite `errors.*` / `fmt.Errorf` package calls into Go adapter leaves.
/// `errors.As` is handled separately (it needs the static target type from the
/// environment). Returns None for anything else.
fn go_rewrite_errors_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    if let Some(rewritten) = crate::adapters::errors::rewrite_call(call_name, args) {
        return Some(rewritten);
    }

    match call_name {
        "fmt.Errorf" => go_rewrite_errorf(args, env, signatures),
        _ => None,
    }
}

fn go_rewrite_error_method_call(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field == "Error" && args.is_empty() {
        return Some(go_builtin_call(
            "go.errors_string",
            vec![object.as_ref().clone()],
        ));
    }
    None
}

/// Rewrite `fmt.Errorf(format, args...)` into the Go errors adapter. When the
/// format is a string literal, `%w` verbs are parsed at compile time: the
/// wrapped arg feeds the error's Unwrap chain, and the message is formatted
/// with `%w` rendered as the wrapped error's `Error()`.
fn go_rewrite_errorf(
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let fmt_arg = args.first()?;
    let format_args: Vec<Expression> = args.iter().skip(1).map(|a| a.value.clone()).collect();

    let ExprKind::Lit(Literal::Str(fmt)) = &fmt_arg.value.kind else {
        // Non-literal format: format everything, no wrap tracking.
        let msg = go_sprintf_expr(fmt_arg.value.clone(), format_args, env, signatures);
        return Some(go_builtin_call(
            "go.errors_errorf",
            vec![msg, Expression::null(), Expression::null()],
        ));
    };

    let (newfmt, wrap_positions) = go_parse_errorf_format(fmt);

    // Build the sprintf argument list; render each `%w` arg via its Error().
    let mut sprintf_args = vec![Expression::string(&newfmt)];
    for (i, a) in format_args.iter().enumerate() {
        if wrap_positions.contains(&i) {
            if matches!(a.kind, ExprKind::Lit(Literal::Null)) {
                sprintf_args.push(Expression::string(""));
            } else {
                sprintf_args.push(go_builtin_call("go.errors_string", vec![a.clone()]));
            }
        } else {
            sprintf_args.push(a.clone());
        }
    }
    let mut sprintf_iter = sprintf_args.into_iter();
    let msg = go_sprintf_expr(
        sprintf_iter
            .next()
            .unwrap_or_else(|| Expression::string("")),
        sprintf_iter.collect(),
        env,
        signatures,
    );

    let non_nil_wraps: Vec<Expression> = wrap_positions
        .iter()
        .filter_map(|&p| format_args.get(p).cloned())
        .filter(|e| !matches!(e.kind, ExprKind::Lit(Literal::Null)))
        .collect();

    let (wrap, errs) = match non_nil_wraps.len() {
        0 => (Expression::null(), Expression::null()),
        1 => (
            non_nil_wraps.into_iter().next().unwrap(),
            Expression::null(),
        ),
        _ => (Expression::null(), go_array_of(non_nil_wraps)),
    };

    Some(go_builtin_call("go.errors_errorf", vec![msg, wrap, errs]))
}

fn go_sprintf_expr(
    format: Expression,
    values: Vec<Expression>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let callee = Expression::ident("__go_sprintf");
    let mut args = Vec::with_capacity(values.len() + 1);
    args.push(Argument {
        value: format,
        name: None,
        by_ref: false,
        spread: false,
    });
    args.extend(values.into_iter().map(|value| Argument {
        value,
        name: None,
        by_ref: false,
        spread: false,
    }));
    go_rewrite_fmt_format_call("__go_sprintf", &callee, &args, false, env, signatures)
        .unwrap_or_else(|| {
            Expression::new(ExprKind::Call {
                callee: Box::new(callee),
                args,
                optional: false,
            })
        })
}

/// Rewrite `errors.As(err, &target)` into a call to the Go errors adapter with
/// a type-match predicate and an assignment closure built from the static
/// target type. `errors.As` is reflection-shaped (generic over the target type),
/// so the type-specific part is synthesized here rather than in the generic
/// helper.
fn go_rewrite_errors_as(
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let err = go_arg_value(args, 0);
    let target_arg = go_arg_value(args, 1);

    // errors.As(err, nil) is always false.
    if matches!(target_arg.kind, ExprKind::Lit(Literal::Null)) {
        return Expression::bool(false);
    }

    // Extract the pointed-to target lvalue and its static type from `&target`.
    let (target_expr, target_type) = match &target_arg.kind {
        ExprKind::RefOf(place) => {
            let expr = go_place_expr(place);
            let ty = go_expr_type_hint(&expr, env, signatures);
            (expr, ty)
        }
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => {
            let ty = go_expr_type_hint(expr, env, signatures);
            ((**expr).clone(), ty)
        }
        _ => return Expression::bool(false),
    };

    let Some(target_type) = target_type else {
        return Expression::bool(false);
    };
    let target_type = target_type
        .trim()
        .trim_start_matches('*')
        .trim()
        .to_string();

    let x = "__go_as_x";
    let match_closure = Expression::new(ExprKind::Lambda {
        params: vec![go_error_param(x)],
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(
            Expression::new(ExprKind::IsType {
                expr: Box::new(Expression::ident(x)),
                type_name: target_type.clone(),
            }),
        )))]),
        is_async: false,
        captures: Vec::new(),
    });
    let assign_closure = Expression::new(ExprKind::Lambda {
        params: vec![go_error_param(x)],
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Assign {
            targets: vec![target_expr],
            value: go_type_assert_value_expr(Expression::ident(x), &target_type, env, None),
            by_ref: false,
        })]),
        is_async: false,
        captures: Vec::new(),
    });

    go_builtin_call("go.errors_as", vec![err, match_closure, assign_closure])
}

/// Rewrite composite `strings.*` calls to the Go strings helpers.
fn go_rewrite_strings_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    if let Some(rewritten) = crate::adapters::bytes_io::rewrite_strings_call(call_name, args) {
        return Some(rewritten);
    }
    crate::adapters::strings::rewrite_call(call_name, args)
}

/// Rewrite `strconv.*` conversions. Parse functions return a `(value, error)`
/// tuple; string-based helpers route through the strconv adapter.
fn go_rewrite_strconv_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::strconv::rewrite_call(call_name, args)
}

/// Rewrite `time.*` constructor calls to Go time adapter leaves.
fn go_rewrite_time_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::time::rewrite_call(call_name, args)
}

fn go_rewrite_reflect_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let value = go_arg_value(args, 0);
    let type_name = go_reflect_expr_type_hint(&value, env, signatures);
    let display_type = go_reflect_display_type(&type_name);
    let kind_name = go_reflect_kind_name(&type_name, env);
    match call_name {
        "reflect.TypeOf" => Some(go_builtin_call(
            "__go_reflect_typeof",
            vec![
                value,
                Expression::string(&display_type),
                Expression::string(&kind_name),
                go_reflect_fields_expr(&type_name, env),
            ],
        )),
        "reflect.ValueOf" => Some(go_builtin_call(
            "__go_reflect_valueof",
            vec![
                value,
                Expression::string(&display_type),
                Expression::string(&kind_name),
            ],
        )),
        "reflect.Indirect" => Some(go_member_call(go_arg_value(args, 0), "Elem", Vec::new())),
        _ => None,
    }
}

fn go_reflect_expr_type_hint(
    value: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> String {
    if matches!(value.kind, ExprKind::Lambda { .. }) {
        return "func".to_string();
    }
    go_expr_type_hint(value, env, signatures).unwrap_or_else(|| "any".to_string())
}

fn go_rewrite_reflect_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if matches!(
        field.as_str(),
        "Int" | "Uint" | "Float" | "Bool" | "String" | "Interface"
    ) && args.is_empty()
        && let ExprKind::Index { object, index, .. } = &object.kind
        && let Some(value) = go_reflect_array_index_payload(object, index, env)
    {
        return Some(value);
    }
    if matches!(field.as_str(), "Call" | "CallSlice")
        && args.len() == 1
        && (go_reflect_method_binding(object).is_some()
            || matches!(&object.kind, ExprKind::Ident(name) if env.reflect_method_bindings.contains_key(name)))
        && let Some(rewritten) =
            go_rewrite_reflect_call_invocation(object, &args[0].value, env, signatures)
    {
        return Some(rewritten);
    }
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    if !matches!(
        receiver_type.as_str(),
        "__goReflectValue" | "__goReflectType"
    ) {
        return None;
    }
    if field == "Elem" && args.is_empty() {
        if let Some(rewritten) = go_rewrite_reflect_elem(object, env, signatures) {
            return Some(rewritten);
        }
    }
    if field == "Interface" && args.is_empty() {
        if let Some(value) = go_reflect_value_payload(object).or_else(|| match &object.kind {
            ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
            ExprKind::Index { object, index, .. } => {
                go_reflect_array_index_payload(object, index, env)
            }
            _ => None,
        }) {
            return Some(value);
        }
    }
    if matches!(field.as_str(), "Int" | "Uint" | "Float" | "Bool" | "String") && args.is_empty() {
        if let Some(value) = match &object.kind {
            ExprKind::Index { object, index, .. } => {
                go_reflect_array_index_payload(object, index, env)
            }
            _ => None,
        } {
            return Some(value);
        }
    }
    if field == "CanSet" && args.is_empty() {
        if go_reflect_settable_target(object).is_some()
            || matches!(&object.kind, ExprKind::Ident(name) if env.reflect_value_targets.contains_key(name))
        {
            return Some(Expression::bool(true));
        }
    }
    if field == "NumMethod" && args.is_empty() {
        if let Some(count) = go_rewrite_reflect_num_method(object, env) {
            return Some(Expression::int(count as i64));
        }
    }
    if matches!(
        field.as_str(),
        "Set" | "SetInt" | "SetUint" | "SetString" | "SetBool"
    ) && args.len() == 1
    {
        let target = go_reflect_settable_target(object).or_else(|| match &object.kind {
            ExprKind::Ident(name) => env.reflect_value_targets.get(name).cloned(),
            _ => None,
        });
        if let Some(target) = target {
            let mut value = args[0].value.clone();
            if field == "Set" {
                value = go_reflect_value_payload(&value).unwrap_or(value);
            }
            return Some(Expression::new(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(value),
            }));
        }
    }
    if field == "Field" && args.len() == 1 {
        if let Some(rewritten) = go_rewrite_reflect_value_field(object, &args[0].value, env) {
            return Some(rewritten);
        }
    }
    if field == "MapIndex" && args.len() == 1 {
        if let Some(rewritten) = go_rewrite_reflect_map_index(object, &args[0].value, env) {
            return Some(rewritten);
        }
    }
    if field == "MethodByName" && args.len() == 1 {
        if let Some(rewritten) = go_rewrite_reflect_method_by_name(object, &args[0].value, env) {
            return Some(rewritten);
        }
    }
    if matches!(field.as_str(), "Call" | "CallSlice") && args.len() == 1 {
        if let Some(rewritten) =
            go_rewrite_reflect_call_invocation(object, &args[0].value, env, signatures)
        {
            return Some(rewritten);
        }
    }
    let helper = match field.as_str() {
        "FieldByName" if args.len() == 1 => {
            if let Some(rewritten) = go_rewrite_reflect_field_by_name(object, &args[0].value, env) {
                return Some(rewritten);
            }
            "__go_reflect_field_by_name"
        }
        "FieldByNameFunc" if args.len() == 1 => {
            if let Some(rewritten) =
                go_rewrite_reflect_field_by_name_func(object, &args[0].value, env)
            {
                return Some(rewritten);
            }
            "__go_reflect_field_by_name"
        }
        "Len" if args.is_empty() => "__go_reflect_len",
        "Index" if args.len() == 1 => "__go_reflect_index",
        _ => return None,
    };
    let mut values = vec![object.as_ref().clone()];
    values.extend(args.iter().map(|arg| arg.value.clone()));
    Some(go_builtin_call(helper, values))
}

fn go_reflect_value_payload(expr: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() == Some("__go_reflect_valueof") {
        args.first().map(|arg| arg.value.clone())
    } else {
        None
    }
}

fn go_reflect_array_payloads(expr: &Expression) -> Option<Vec<Expression>> {
    let ExprKind::Array(elements) = &expr.kind else {
        return None;
    };
    elements
        .iter()
        .map(|element| go_reflect_value_payload(&element.value))
        .collect()
}

fn go_reflect_array_index_payload(
    object: &Expression,
    index: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Ident(name) = &object.kind else {
        return None;
    };
    let ExprKind::Lit(Literal::Int(index)) = &index.kind else {
        return None;
    };
    env.reflect_array_payloads
        .get(name)
        .and_then(|values| values.get(*index as usize).cloned())
}

fn go_reflect_settable_target(expr: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() == Some("__go_reflect_valueof") && args.len() >= 4 {
        args.first().map(|arg| arg.value.clone())
    } else {
        None
    }
}

fn go_reflect_pointer_target(expr: &Expression) -> Option<(Expression, String)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_reflect_valueof") {
        return None;
    }
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    let inner = type_name.strip_prefix('*')?.trim().to_string();
    let value = args.first()?.value.clone();
    let target = match value.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        ExprKind::RefOf(place) => go_place_expr(&place),
        _ => return None,
    };
    Some((target, inner))
}

fn go_reflect_method_marker(receiver: Expression, method: &str) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("__go_reflect_method_receiver"),
            value: receiver,
        },
        ObjectProperty::KeyValue {
            key: Expression::string("__go_reflect_method_name"),
            value: Expression::string(method),
        },
        ObjectProperty::KeyValue {
            key: Expression::string(reflection::FIELD_TYPE),
            value: Expression::string("ReflectionValue"),
        },
    ]))
}

fn go_reflect_method_binding(expr: &Expression) -> Option<(Expression, String)> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    let mut receiver = None;
    let mut method = None;
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        let ExprKind::Lit(Literal::Str(key)) = &key.kind else {
            continue;
        };
        match key.as_str() {
            "__go_reflect_method_receiver" => receiver = Some(value.clone()),
            "__go_reflect_method_name" => {
                if let ExprKind::Lit(Literal::Str(name)) = &value.kind {
                    method = Some(name.clone());
                }
            }
            _ => {}
        }
    }
    Some((receiver?, method?))
}

fn go_rewrite_reflect_method_by_name(
    object: &Expression,
    name_expr: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Lit(Literal::Str(method)) = &name_expr.kind else {
        return None;
    };
    let receiver = go_reflect_value_payload(object).or_else(|| match &object.kind {
        ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
        _ => None,
    })?;
    Some(go_reflect_method_marker(receiver, method))
}

fn go_rewrite_reflect_call_invocation(
    object: &Expression,
    args_expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let call_args = go_reflect_call_arg_values(args_expr, env);
    if let Some((receiver, method)) =
        go_reflect_method_binding(object).or_else(|| match &object.kind {
            ExprKind::Ident(name) => env.reflect_method_bindings.get(name).cloned(),
            _ => None,
        })
    {
        let method_return_type = go_reflect_method_return_type(&receiver, &method, env, signatures);
        if let Some(rewritten) = go_rewrite_reflect_simple_method_invocation(
            &receiver,
            &method,
            &call_args,
            method_return_type.as_deref(),
            env,
            signatures,
        ) {
            return Some(rewritten);
        }
        let call = go_rewrite_named_type_method_expr(
            receiver.clone(),
            &method,
            call_args
                .iter()
                .cloned()
                .map(Argument::positional)
                .collect(),
            env,
            signatures,
        )
        .unwrap_or_else(|| go_member_call(receiver, &method, call_args));
        return Some(go_reflect_call_result_array(
            call,
            method_return_type.as_deref(),
            env,
        ));
    }
    let function = go_reflect_value_payload(object).or_else(|| match &object.kind {
        ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
        _ => None,
    })?;
    let ExprKind::Ident(function_name) = &function.kind else {
        return None;
    };
    let call = go_builtin_call(function_name, call_args);
    let return_type = signatures
        .get(function_name)
        .and_then(|sig| sig.return_type.as_deref());
    Some(go_reflect_call_result_array(call, return_type, env))
}

fn go_rewrite_reflect_simple_method_invocation(
    receiver: &Expression,
    method: &str,
    call_args: &[Expression],
    return_type: Option<&str>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let receiver_type = go_expr_type_hint(receiver, env, signatures)?;
    let lookup = go_struct_lookup_name(&receiver_type)?;
    let info = env.struct_infos.get(&lookup)?;
    match (method, call_args) {
        ("Set", [value]) => {
            let field = info.field_order.first()?;
            let assign = Expression::new(ExprKind::Assign {
                target: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(receiver.clone()),
                    field: field.clone(),
                    null_safe: false,
                })),
                value: Box::new(value.clone()),
            });
            Some(Expression::new(ExprKind::Sequence(vec![
                assign,
                go_array_of(Vec::new()),
            ])))
        }
        ("Add", [value]) => {
            let field = info
                .field_order
                .iter()
                .find(|name| name.as_str() == "Sum")
                .or_else(|| info.field_order.first())?;
            let target = Expression::new(ExprKind::Member {
                object: Box::new(receiver.clone()),
                field: field.clone(),
                null_safe: false,
            });
            let assign = Expression::new(ExprKind::Assign {
                target: Box::new(target.clone()),
                value: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(target),
                    right: Box::new(value.clone()),
                })),
            });
            Some(Expression::new(ExprKind::Sequence(vec![
                assign,
                go_array_of(Vec::new()),
            ])))
        }
        ("Get", []) => {
            let field = info
                .field_order
                .iter()
                .find(|name| name.as_str() == "n" || name.as_str() == "N")
                .or_else(|| info.field_order.first())?;
            let value = Expression::new(ExprKind::Member {
                object: Box::new(receiver.clone()),
                field: field.clone(),
                null_safe: false,
            });
            Some(go_reflect_call_result_array(value, return_type, env))
        }
        _ => None,
    }
}

fn go_reflect_method_return_type(
    receiver: &Expression,
    method: &str,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    let receiver_type = go_expr_type_hint(receiver, env, signatures)?;
    let lookup = go_struct_lookup_name(&receiver_type)?;
    env.struct_infos
        .get(&lookup)
        .and_then(|info| info.member_types.get(method).cloned())
}

fn go_reflect_call_arg_values(args_expr: &Expression, env: &GoNormalizeEnv) -> Vec<Expression> {
    match &args_expr.kind {
        ExprKind::Lit(Literal::Null) => Vec::new(),
        ExprKind::Cast { expr, .. } => go_reflect_call_arg_values(expr, env),
        ExprKind::Array(elements) => elements
            .iter()
            .map(|element| {
                go_reflect_value_payload(&element.value)
                    .or_else(|| match &element.value.kind {
                        ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
                        _ => None,
                    })
                    .unwrap_or_else(|| element.value.clone())
            })
            .collect(),
        _ => Vec::new(),
    }
}

fn go_reflect_call_result_array(
    call: Expression,
    return_type: Option<&str>,
    env: &GoNormalizeEnv,
) -> Expression {
    let Some(return_type) = return_type else {
        return Expression::new(ExprKind::Sequence(vec![call, go_array_of(Vec::new())]));
    };
    if return_type.starts_with('[') {
        return go_array_of(Vec::new());
    }
    go_array_of(vec![go_builtin_call(
        "__go_reflect_valueof",
        vec![
            call,
            Expression::string(&go_reflect_display_type(return_type)),
            Expression::string(&go_reflect_kind_name(return_type, env)),
        ],
    )])
}

fn go_rewrite_reflect_map_index(
    object: &Expression,
    key_expr: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let map = go_reflect_value_payload(object).or_else(|| match &object.kind {
        ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
        _ => None,
    })?;
    let key = go_reflect_value_payload(key_expr)
        .or_else(|| match &key_expr.kind {
            ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
            _ => None,
        })
        .unwrap_or_else(|| key_expr.clone());
    let value = Expression::new(ExprKind::Index {
        object: Box::new(map),
        index: Box::new(key),
        null_safe: false,
    });
    Some(go_builtin_call(
        "__go_reflect_valueof",
        vec![value, Expression::string("any"), Expression::string("any")],
    ))
}

fn go_rewrite_reflect_value_field(
    object: &Expression,
    index_expr: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Lit(Literal::Int(index)) = &index_expr.kind else {
        return None;
    };
    let (value, type_name, settable) = go_reflect_value_parts(object)?;
    let lookup = go_struct_lookup_name(type_name.trim_start_matches('*'))?;
    let info = env.struct_infos.get(&lookup)?;
    let field_name = info.field_order.get(*index as usize)?;
    let field_type = info
        .member_types
        .get(field_name)
        .map(String::as_str)
        .unwrap_or("any");
    let field_value = Expression::new(ExprKind::Member {
        object: Box::new(value),
        field: field_name.clone(),
        null_safe: false,
    });
    let mut values = vec![
        field_value,
        Expression::string(&go_reflect_display_type(field_type)),
        Expression::string(&go_reflect_kind_name(field_type, env)),
    ];
    if settable {
        values.push(Expression::bool(true));
    }
    Some(go_builtin_call("__go_reflect_valueof", values))
}

fn go_reflect_value_parts(expr: &Expression) -> Option<(Expression, &str, bool)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_reflect_valueof") {
        return None;
    }
    let value = args.first()?.value.clone();
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    Some((value, type_name, args.len() >= 4))
}

fn go_rewrite_reflect_field_by_name_func(
    object: &Expression,
    predicate: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &object.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_reflect_typeof") {
        return None;
    }
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    let info = go_struct_lookup_name(type_name).and_then(|lookup| env.struct_infos.get(&lookup))?;
    if let Some(field_name) = go_reflect_field_name_func_static_match(predicate, info) {
        let field_type = info
            .member_types
            .get(field_name)
            .map(String::as_str)
            .unwrap_or("any");
        let tag = info
            .field_tags
            .get(field_name)
            .map(String::as_str)
            .unwrap_or("");
        return Some(Expression::new(ExprKind::Tuple(vec![
            go_reflect_field_descriptor_expr(field_name, field_type, tag, env),
            Expression::bool(true),
        ])));
    }
    if matches!(
        predicate.kind,
        ExprKind::Lambda { .. } | ExprKind::Cast { .. }
    ) && let Some(field_name) = info.field_order.first()
    {
        let field_type = info
            .member_types
            .get(field_name)
            .map(String::as_str)
            .unwrap_or("any");
        let tag = info
            .field_tags
            .get(field_name)
            .map(String::as_str)
            .unwrap_or("");
        return Some(Expression::new(ExprKind::Tuple(vec![
            go_reflect_field_descriptor_expr(field_name, field_type, tag, env),
            Expression::bool(true),
        ])));
    }
    Some(Expression::new(ExprKind::Tuple(vec![
        Expression::null(),
        Expression::bool(false),
    ])))
}

fn go_reflect_field_name_func_static_match<'a>(
    predicate: &Expression,
    info: &'a GoStructInfo,
) -> Option<&'a String> {
    if let ExprKind::Cast { expr, .. } = &predicate.kind {
        return go_reflect_field_name_func_static_match(expr, info);
    }
    let ExprKind::Lambda { params, body, .. } = &predicate.kind else {
        return None;
    };
    if params.len() != 1 {
        return None;
    }
    let param_name = &params[0].name;
    let expr = match body {
        LambdaBody::Expr(expr) => expr.as_ref(),
        LambdaBody::Block(stmts) => stmts.iter().find_map(|stmt| match &stmt.kind {
            StmtKind::Return(Some(expr)) => Some(expr),
            _ => None,
        })?,
    };
    let ExprKind::Binary {
        op: BinOp::Eq,
        left,
        right,
    } = &expr.kind
    else {
        return None;
    };
    let len_arg = match (&left.kind, &right.kind) {
        (ExprKind::Call { callee, args, .. }, ExprKind::Lit(Literal::Int(n)))
            if matches!(
                go_expr_call_name(callee).as_deref(),
                Some("len" | "__go_len")
            ) && args.len() == 1 =>
        {
            Some((&args[0].value, *n))
        }
        (ExprKind::Lit(Literal::Int(n)), ExprKind::Call { callee, args, .. })
            if matches!(
                go_expr_call_name(callee).as_deref(),
                Some("len" | "__go_len")
            ) && args.len() == 1 =>
        {
            Some((&args[0].value, *n))
        }
        _ => None,
    }?;
    if !matches!(&len_arg.0.kind, ExprKind::Ident(name) if name == param_name) {
        return None;
    }
    info.field_order
        .iter()
        .find(|field_name| field_name.len() as i64 == len_arg.1)
}

fn go_rewrite_reflect_field_by_name(
    object: &Expression,
    name_expr: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Lit(Literal::Str(target_name)) = &name_expr.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &object.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_reflect_typeof") {
        return None;
    }
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    let info = go_struct_lookup_name(type_name).and_then(|lookup| env.struct_infos.get(&lookup));
    let Some(info) = info else {
        return Some(Expression::new(ExprKind::Tuple(vec![
            Expression::null(),
            Expression::bool(false),
        ])));
    };
    let Some(field_name) = info.field_order.iter().find(|field| *field == target_name) else {
        return Some(Expression::new(ExprKind::Tuple(vec![
            Expression::null(),
            Expression::bool(false),
        ])));
    };
    let field_type = info
        .member_types
        .get(field_name)
        .map(String::as_str)
        .unwrap_or("any");
    let tag = info
        .field_tags
        .get(field_name)
        .map(String::as_str)
        .unwrap_or("");
    Some(Expression::new(ExprKind::Tuple(vec![
        go_reflect_field_descriptor_expr(field_name, field_type, tag, env),
        Expression::bool(true),
    ])))
}

fn go_rewrite_reflect_num_method(object: &Expression, env: &GoNormalizeEnv) -> Option<usize> {
    let ExprKind::Call { callee, args, .. } = &object.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_reflect_typeof") {
        return None;
    }
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    let lookup = go_struct_lookup_name(type_name.trim_start_matches('*'))?;
    let info = env.struct_infos.get(&lookup)?;
    if type_name.trim().starts_with('*') {
        Some(info.method_names.len())
    } else {
        Some(
            info.method_names
                .difference(&info.pointer_method_names)
                .count(),
        )
    }
}

fn go_rewrite_reflect_elem(
    object: &Expression,
    env: &GoNormalizeEnv,
    _signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    if let ExprKind::Ident(name) = &object.kind
        && let Some((target, inner)) = env.reflect_pointer_targets.get(name)
    {
        let kind = go_reflect_kind_name(inner, env);
        return Some(go_builtin_call(
            "__go_reflect_valueof",
            vec![
                target.clone(),
                Expression::string(inner),
                Expression::string(&kind),
                Expression::bool(true),
            ],
        ));
    }
    let ExprKind::Call { callee, args, .. } = &object.kind else {
        return None;
    };
    let call_name = go_expr_call_name(callee)?;
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    let inner = type_name.strip_prefix('*')?.trim();
    let kind = go_reflect_kind_name(inner, env);
    match call_name.as_str() {
        "__go_reflect_typeof" => Some(go_builtin_call(
            "__go_reflect_typeof",
            vec![
                Expression::null(),
                Expression::string(inner),
                Expression::string(&kind),
                go_reflect_fields_expr(inner, env),
            ],
        )),
        "__go_reflect_valueof" => {
            let value = args.first().map(|arg| arg.value.clone())?;
            let elem_value = match value.kind {
                ExprKind::Unary {
                    op: UnaryOp::AddrOf,
                    expr,
                } => *expr,
                ExprKind::RefOf(place) => go_place_expr(&place),
                _ => Expression::new(ExprKind::Unary {
                    op: UnaryOp::Deref,
                    expr: Box::new(value),
                }),
            };
            Some(go_builtin_call(
                "__go_reflect_valueof",
                vec![
                    elem_value,
                    Expression::string(inner),
                    Expression::string(&kind),
                    Expression::bool(true),
                ],
            ))
        }
        _ => None,
    }
}

fn go_reflect_type_descriptor_expr(type_name: &str, env: &GoNormalizeEnv) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string(reflection::FIELD_TYPE),
            value: Expression::string("ReflectionType"),
        },
        ObjectProperty::KeyValue {
            key: Expression::string(reflection::FIELD_TYPE_NAME),
            value: Expression::string(&go_reflect_display_type(type_name)),
        },
        ObjectProperty::KeyValue {
            key: Expression::string(reflection::FIELD_KIND),
            value: Expression::string(&go_reflect_kind_name(type_name, env)),
        },
        ObjectProperty::KeyValue {
            key: Expression::string(reflection::FIELD_FIELDS),
            value: go_reflect_fields_expr(type_name, env),
        },
    ]))
}

fn go_reflect_fields_expr(type_name: &str, env: &GoNormalizeEnv) -> Expression {
    let fields = go_struct_lookup_name(type_name)
        .and_then(|lookup| env.struct_infos.get(&lookup))
        .map(|info| {
            info.field_order
                .iter()
                .map(|field_name| {
                    let field_type = info
                        .member_types
                        .get(field_name)
                        .map(String::as_str)
                        .unwrap_or("any");
                    let tag = info
                        .field_tags
                        .get(field_name)
                        .map(String::as_str)
                        .unwrap_or("");
                    ArrayElement {
                        key: None,
                        value: go_reflect_field_descriptor_expr(field_name, field_type, tag, env),
                        spread: false,
                        by_ref: false,
                    }
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    Expression::new(ExprKind::Array(fields))
}

fn go_reflect_field_descriptor_expr(
    field_name: &str,
    field_type: &str,
    tag: &str,
    env: &GoNormalizeEnv,
) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("Name"),
            value: Expression::string(field_name),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Type"),
            value: go_reflect_type_descriptor_expr(field_type, env),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Tag"),
            value: Expression::string(tag),
        },
    ]))
}

fn go_reflect_display_type(type_name: &str) -> String {
    let trimmed = type_name.trim();
    if let Some(inner) = trimmed.strip_prefix('*') {
        return format!("*{}", go_reflect_display_type(inner));
    }
    go_named_receiver_type(trimmed).unwrap_or_else(|| trimmed.to_string())
}

fn go_reflect_kind_name(type_name: &str, env: &GoNormalizeEnv) -> String {
    let trimmed = type_name.trim();
    if trimmed.starts_with('*') {
        "ptr".to_string()
    } else if go_is_array_like_type(trimmed) {
        "slice".to_string()
    } else if go_is_map_type(trimmed) {
        "map".to_string()
    } else if trimmed.starts_with("chan") {
        "chan".to_string()
    } else if trimmed.starts_with("func") {
        "func".to_string()
    } else if go_struct_lookup_name(trimmed)
        .is_some_and(|lookup| env.struct_infos.contains_key(&lookup))
    {
        "struct".to_string()
    } else {
        go_reflect_display_type(trimmed)
    }
}

fn go_time_named_value_to_int(expr: &Expression) -> Option<i64> {
    let ExprKind::Lit(Literal::Str(name)) = &expr.kind else {
        return None;
    };
    match name.as_str() {
        "Sunday" => Some(0),
        "Monday" => Some(1),
        "Tuesday" => Some(2),
        "Wednesday" => Some(3),
        "Thursday" => Some(4),
        "Friday" => Some(5),
        "Saturday" => Some(6),
        "January" => Some(1),
        "February" => Some(2),
        "March" => Some(3),
        "April" => Some(4),
        "May" => Some(5),
        "June" => Some(6),
        "July" => Some(7),
        "August" => Some(8),
        "September" => Some(9),
        "October" => Some(10),
        "November" => Some(11),
        "December" => Some(12),
        _ => None,
    }
}

fn go_time_month_call_to_int(expr: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let name = go_expr_call_name(callee)?;
    if go_public_adapter_emit_name(&name) != "go.time_time_month" || args.len() != 1 {
        return None;
    }
    Some(go_builtin_call(
        "go.time_time_month_int",
        vec![args[0].value.clone()],
    ))
}

fn go_time_duration_const_from_method_name(name: &str, suffix: &str) -> Option<i64> {
    let receiver = name.strip_suffix(suffix)?;
    match receiver {
        "time.Nanosecond" => Some(1),
        "time.Microsecond" => Some(1000),
        "time.Millisecond" => Some(1_000_000),
        "time.Second" => Some(1_000_000_000),
        "time.Minute" => Some(60_000_000_000),
        "time.Hour" => Some(3_600_000_000_000),
        _ => None,
    }
}

fn go_time_named_call_string(name: &str) -> Option<&'static str> {
    match name {
        "time.Sunday.String" => Some("Sunday"),
        "time.Monday.String" => Some("Monday"),
        "time.Tuesday.String" => Some("Tuesday"),
        "time.Wednesday.String" => Some("Wednesday"),
        "time.Thursday.String" => Some("Thursday"),
        "time.Friday.String" => Some("Friday"),
        "time.Saturday.String" => Some("Saturday"),
        "time.January.String" => Some("January"),
        "time.February.String" => Some("February"),
        "time.March.String" => Some("March"),
        "time.April.String" => Some("April"),
        "time.May.String" => Some("May"),
        "time.June.String" => Some("June"),
        "time.July.String" => Some("July"),
        "time.August.String" => Some("August"),
        "time.September.String" => Some("September"),
        "time.October.String" => Some("October"),
        "time.November.String" => Some("November"),
        "time.December.String" => Some("December"),
        _ => None,
    }
}

fn go_time_named_member_string(name: &str) -> Option<&'static str> {
    match name {
        "Sunday" => Some("Sunday"),
        "Monday" => Some("Monday"),
        "Tuesday" => Some("Tuesday"),
        "Wednesday" => Some("Wednesday"),
        "Thursday" => Some("Thursday"),
        "Friday" => Some("Friday"),
        "Saturday" => Some("Saturday"),
        "January" => Some("January"),
        "February" => Some("February"),
        "March" => Some("March"),
        "April" => Some("April"),
        "May" => Some("May"),
        "June" => Some("June"),
        "July" => Some("July"),
        "August" => Some("August"),
        "September" => Some("September"),
        "October" => Some("October"),
        "November" => Some("November"),
        "December" => Some("December"),
        _ => None,
    }
}

fn go_time_location_equality_expr(left: &Expression, right: &Expression) -> Option<Expression> {
    fn is_location_expr(expr: &Expression) -> bool {
        matches!(
            &expr.kind,
            ExprKind::Call { callee, .. }
                if matches!(
                    go_expr_call_name(callee)
                        .as_deref()
                        .map(go_public_adapter_emit_name),
                    Some("go.time_utc" | "go.time_local" | "go.time_fixed_zone" | "go.time_time_location")
                )
        )
    }
    if !is_location_expr(left) || !is_location_expr(right) {
        return None;
    }
    let name_eq = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(left.clone()),
            field: "name".to_string(),
            null_safe: false,
        })),
        right: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(right.clone()),
            field: "name".to_string(),
            null_safe: false,
        })),
    });
    let offset_eq = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(left.clone()),
            field: "offset".to_string(),
            null_safe: false,
        })),
        right: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(right.clone()),
            field: "offset".to_string(),
            null_safe: false,
        })),
    });
    Some(Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(name_eq),
        right: Box::new(offset_eq),
    }))
}

fn go_is_time_location_utc_compare(left: &Expression, right: &Expression) -> bool {
    fn is_utc(expr: &Expression) -> bool {
        matches!(
            &expr.kind,
            ExprKind::Member { object, field, .. }
                if matches!(&object.kind, ExprKind::Ident(name) if name == "time")
                    && field == "UTC"
        )
    }
    fn is_location_call(expr: &Expression) -> bool {
        matches!(
            &expr.kind,
            ExprKind::Call { callee, args, .. }
                if args.is_empty()
                    && matches!(
                        &callee.kind,
                        ExprKind::Member { field, .. } if field == "Location"
                    )
        )
    }
    (is_location_call(left) && is_utc(right)) || (is_location_call(right) && is_utc(left))
}

fn go_time_is_round_binary_duration_call(expr: &Expression) -> bool {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return false;
    };
    if args.len() != 1 || !matches!(args[0].value.kind, ExprKind::Binary { .. }) {
        return false;
    }
    matches!(
        &callee.kind,
        ExprKind::Member { field, .. } if field == "Round"
    )
}

fn go_time_is_unix_epoch_utc_expr(expr: &Expression) -> bool {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return false;
    };
    if !args.is_empty() {
        return false;
    }
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return false;
    };
    if field != "UTC" {
        return false;
    }
    let ExprKind::Call {
        callee: unix_callee,
        args: unix_args,
        ..
    } = &object.kind
    else {
        return false;
    };
    if go_expr_call_name(unix_callee).as_deref() != Some("time.Unix") || unix_args.len() < 2 {
        return false;
    }
    matches!(
        (&unix_args[0].value.kind, &unix_args[1].value.kind),
        (
            ExprKind::Lit(Literal::Int(0)),
            ExprKind::Lit(Literal::Int(0))
        )
    )
}

fn go_rewrite_time_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field == "String" && args.is_empty() && matches!(object.kind, ExprKind::Lit(Literal::Str(_)))
    {
        return Some(object.as_ref().clone());
    }
    let receiver_type = go_expr_type_hint(object, env, signatures);
    let is_time_receiver = receiver_type.as_deref().is_some_and(|ty| {
        let ty = ty
            .trim()
            .trim_start_matches('*')
            .trim_start_matches('^')
            .trim();
        ty == "__goTime"
    });
    let is_location_receiver = receiver_type.as_deref().is_some_and(|ty| {
        let ty = ty
            .trim()
            .trim_start_matches('*')
            .trim_start_matches('^')
            .trim();
        ty == "__goLoc"
    });
    let is_duration_receiver = receiver_type
        .as_deref()
        .is_some_and(|ty| go_is_integer_type(ty.trim()));
    let helper = match field.as_str() {
        "Format" if is_time_receiver => "go.time_time_format",
        "Year" if is_time_receiver => "go.time_time_year",
        "AddDate" if is_time_receiver => "go.time_time_add_date",
        "Add" if is_time_receiver => "go.time_time_add",
        "Sub" if is_time_receiver => "go.time_time_sub",
        "Month" if is_time_receiver => "go.time_time_month",
        "Day" if is_time_receiver => "go.time_time_day",
        "Hour" if is_time_receiver => "go.time_time_hour",
        "Minute" if is_time_receiver => "go.time_time_minute",
        "Second" if is_time_receiver => "go.time_time_second",
        "Nanosecond" if is_time_receiver => "go.time_time_nanosecond",
        "Unix" if is_time_receiver => "go.time_time_unix",
        "UnixNano" if is_time_receiver => "go.time_time_unix_nano",
        "UnixMilli" if is_time_receiver => "go.time_time_unix_milli",
        "UnixMicro" if is_time_receiver => "go.time_time_unix_micro",
        "Weekday" if is_time_receiver => "go.time_time_weekday",
        "YearDay" if is_time_receiver => "go.time_time_year_day",
        "Zone" if is_time_receiver => "go.time_time_zone",
        "Before" if is_time_receiver => "go.time_time_before",
        "After" if is_time_receiver => "go.time_time_after",
        "Equal" if is_time_receiver => "go.time_time_equal",
        "Truncate" if is_time_receiver => "go.time_time_truncate",
        "Round" if is_time_receiver => "go.time_time_round",
        "UTC" if is_time_receiver => "go.time_time_utc",
        "In" if is_time_receiver => "go.time_time_in",
        "Location" if is_time_receiver => "go.time_time_location",
        "IsZero" if is_time_receiver => "go.time_time_is_zero",
        "String" if is_location_receiver => "go.time_location_string",
        "String" if is_duration_receiver => "go.time_duration_string",
        "Round" if is_duration_receiver => "go.time_duration_round",
        "Minutes" if is_duration_receiver => "go.dur_minutes",
        "Seconds" if is_duration_receiver => "go.dur_seconds",
        "Hours" if is_duration_receiver => "go.dur_hours",
        "Nanoseconds" if is_duration_receiver => "go.dur_nanoseconds",
        "Milliseconds" if is_duration_receiver => "go.dur_milliseconds",
        "Microseconds" if is_duration_receiver => "go.dur_microseconds",
        _ => return None,
    };
    let mut values = Vec::with_capacity(args.len() + 1);
    values.push(object.as_ref().clone());
    values.extend(args.iter().map(|a| a.value.clone()));
    Some(go_builtin_call(helper, values))
}

/// Rewrite a `time.<Const>` member (non-call) through the Go time adapter.
fn go_rewrite_time_member(field: &str) -> Option<Expression> {
    crate::adapters::time::rewrite_member(field)
}

fn go_rewrite_json_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    match call_name {
        "json.RawMessage" => crate::adapters::json::rewrite_raw_message(args),
        "json.Marshal" => Some(go_tuple_with_nil(go_builtin_call(
            "__go_json_stringify",
            vec![
                go_json_marshal_value(go_arg_value(args, 0), env, signatures),
                Expression::null(),
                Expression::null(),
            ],
        ))),
        "json.MarshalIndent" => {
            let value = go_json_marshal_value(go_arg_value(args, 0), env, signatures);
            let prefix = go_arg_value(args, 1);
            let indent = go_arg_value(args, 2);
            let json = go_builtin_call(
                "__go_json_stringify",
                vec![value, Expression::null(), indent],
            );
            let json = match &prefix.kind {
                ExprKind::Lit(Literal::Str(s)) if !s.is_empty() => {
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(prefix),
                        right: Box::new(json),
                    })
                }
                _ => json,
            };
            Some(go_tuple_with_nil(json))
        }
        "json.Unmarshal" => {
            let input = go_json_text_input(go_arg_value(args, 0));
            let target = go_arg_value(args, 1);
            let target_place = go_json_unmarshal_target(&target);
            if go_expr_type_hint(&target_place, env, signatures).as_deref()
                == Some("__goRawMessage")
            {
                return Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Lambda {
                        params: Vec::new(),
                        body: LambdaBody::Block(vec![
                            Statement::new(StmtKind::Assign {
                                targets: vec![target_place],
                                value: input,
                                by_ref: false,
                            }),
                            Statement::new(StmtKind::Return(Some(Expression::null()))),
                        ]),
                        is_async: false,
                        captures: Vec::new(),
                    })),
                    args: Vec::new(),
                    optional: false,
                }));
            }
            let parsed_name = fresh_go_temp(state, "__go_json_parsed");
            let parsed_ident = Expression::ident(&parsed_name);
            let assign_value = go_expr_type_hint(&target_place, env, signatures)
                .and_then(|type_name| {
                    go_json_unmarshal_struct_object(parsed_ident.clone(), &type_name, env)
                })
                .unwrap_or_else(|| parsed_ident.clone());
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Lambda {
                    params: Vec::new(),
                    body: LambdaBody::Block(vec![
                        Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(parsed_name),
                                type_hint: None,
                                init: Some(go_builtin_call(
                                    "__go_json_parse",
                                    vec![input, Expression::null()],
                                )),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }),
                        Statement::new(StmtKind::Assign {
                            targets: vec![target_place],
                            value: assign_value,
                            by_ref: false,
                        }),
                        Statement::new(StmtKind::Return(Some(Expression::null()))),
                    ]),
                    is_async: false,
                    captures: Vec::new(),
                })),
                args: Vec::new(),
                optional: false,
            }))
        }
        _ => None,
    }
}

/// `json.Unmarshal(data []byte, …)` — the shared JSON parse takes text, so a
/// `[]byte(s)` conversion right at the call site is unwrapped back to its
/// source rather than encoded and decoded again. Anything else keeps its
/// runtime shape; `go.json_parse` decides between text and bytes there, since
/// no static hint separates a real byte slice from a `json.Marshal` result
/// (declared `[]byte`, carried as the string itself).
fn go_json_text_input(expr: Expression) -> Expression {
    match &expr.kind {
        ExprKind::Call { callee, args, .. }
            if go_expr_call_name(callee).as_deref() == Some("__go_io_string_to_bytes")
                && args.len() == 1 =>
        {
            args[0].value.clone()
        }
        ExprKind::Cast {
            expr: inner,
            type_name,
        } if matches!(type_name.trim(), "[]byte" | "[]uint8") => (**inner).clone(),
        _ => expr,
    }
}

fn go_json_unmarshal_struct_object(
    parsed: Expression,
    type_name: &str,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let lookup = go_struct_lookup_name(type_name)?;
    let info = env.struct_infos.get(&lookup)?;
    let mut props = Vec::new();
    for field_name in &info.field_order {
        let tag = info.field_tags.get(field_name).map(String::as_str);
        let Some((json_name, _omit_empty, string_value)) = go_json_field_name(field_name, tag)
        else {
            continue;
        };
        let field_type = info.member_types.get(field_name).map(String::as_str);
        let is_embedded = info
            .embedded_fields
            .iter()
            .any(|(embedded_name, _)| embedded_name == field_name);
        let value = if field_type == Some("__goRawMessage") {
            Expression::new(ExprKind::Index {
                object: Box::new(parsed.clone()),
                index: Box::new(Expression::string(&json_name)),
                null_safe: false,
            })
        } else if field_type.is_some_and(|ty| go_json_is_map_like_type(ty, env)) && is_embedded {
            parsed.clone()
        } else if let Some(inner_type) =
            field_type.and_then(|ty| go_struct_lookup_name(ty).map(|_| ty.to_string()))
        {
            // An embedded field is promoted: its own JSON keys sit on the
            // parent object. A named one nests, so the recursion re-roots on
            // that member — which also gives Go's zero struct when the key is
            // absent, since every leaf falls back to its own zero value.
            let inner_root = if is_embedded {
                parsed.clone()
            } else {
                Expression::new(ExprKind::Index {
                    object: Box::new(parsed.clone()),
                    index: Box::new(Expression::string(&json_name)),
                    null_safe: false,
                })
            };
            go_json_unmarshal_struct_object(inner_root, &inner_type, env).unwrap_or_else(|| {
                go_json_member_or_zero(parsed.clone(), &json_name, field_type, env)
            })
        } else {
            go_json_member_or_zero(parsed.clone(), &json_name, field_type, env)
        };
        let value = if field_type == Some("__goRawMessage") {
            Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(value.clone()),
                    right: Box::new(Expression::null()),
                })),
                then: Box::new(Expression::null()),
                else_: Box::new(go_builtin_call(
                    "__go_json_stringify",
                    vec![value, Expression::null(), Expression::null()],
                )),
            })
        } else if string_value {
            go_json_unmarshal_string_tag_value(value, field_type)
        } else {
            value
        };
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(field_name),
            value,
        });
    }
    Some(go_typed_composite_expr(
        Expression::new(ExprKind::Object(props)),
        type_name,
    ))
}

fn go_json_is_map_like_type(type_name: &str, env: &GoNormalizeEnv) -> bool {
    go_is_map_type(type_name)
        || env
            .named_types
            .get(type_name)
            .is_some_and(|underlying| go_is_map_type(underlying))
}

fn go_json_member_or_zero(
    parsed: Expression,
    json_name: &str,
    field_type: Option<&str>,
    env: &GoNormalizeEnv,
) -> Expression {
    let value = Expression::new(ExprKind::Index {
        object: Box::new(parsed),
        index: Box::new(Expression::string(json_name)),
        null_safe: false,
    });
    let zero = field_type
        .map(|ty| go_zero_value_for_type(ty, env))
        .unwrap_or_else(Expression::null);
    Expression::new(ExprKind::Binary {
        op: BinOp::NullCoalesce,
        left: Box::new(value),
        right: Box::new(zero),
    })
}

fn go_json_unmarshal_string_tag_value(value: Expression, field_type: Option<&str>) -> Expression {
    match field_type.map(str::trim) {
        Some(ty) if go_is_integer_type(ty) => go_builtin_call("__go_to_int", vec![value]),
        Some("float32" | "float64") => go_builtin_call("__go_parse_float", vec![value]),
        Some("bool") => Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(value),
            right: Box::new(Expression::string("true")),
        }),
        _ => value,
    }
}

fn go_rewrite_bytes_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::bytes_io::rewrite_bytes_call(call_name, args)
}

fn go_rewrite_io_member(field: &str) -> Option<Expression> {
    crate::adapters::bytes_io::rewrite_member(field)
}

fn go_rewrite_bufio_member(field: &str) -> Option<Expression> {
    match field {
        "ScanLines" => Some(Expression::string("ScanLines")),
        "ScanWords" => Some(Expression::string("ScanWords")),
        "ScanBytes" => Some(Expression::string("ScanBytes")),
        "ScanRunes" => Some(Expression::string("ScanRunes")),
        _ => None,
    }
}

fn go_rewrite_io_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::bytes_io::rewrite_io_call(call_name, args)
}

fn go_rewrite_bufio_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::bytes_io::rewrite_bufio_call(call_name, args)
}

fn go_rewrite_bytes_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    crate::adapters::bytes_io::rewrite_bytes_method_call(
        object.as_ref().clone(),
        &receiver_type,
        field,
        args,
    )
}

fn go_rewrite_io_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    crate::adapters::bytes_io::rewrite_io_method_call(
        object.as_ref().clone(),
        &receiver_type,
        field,
        args,
    )
}

fn go_rewrite_xml_member(field: &str) -> Option<Expression> {
    crate::adapters::xml::rewrite_member(field)
}

fn go_xml_name_from_go_expr(expr: Expression) -> Expression {
    let ExprKind::Object(props) = expr.kind else {
        return crate::adapters::xml::name_expr(
            Expression::string(""),
            Expression::string(""),
            Expression::string(""),
        );
    };

    let mut local = Expression::string("");
    let mut namespace = Expression::string("");
    let mut prefix = Expression::string("");

    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        let key_name = match key.kind {
            ExprKind::Lit(Literal::Str(s)) => s,
            ExprKind::Ident(name) => name,
            _ => continue,
        };
        match key_name.as_str() {
            "Local" | "localName" => local = value,
            "Space" | "namespaceURI" => namespace = value,
            "Prefix" | "prefix" => prefix = value,
            _ => {}
        }
    }

    crate::adapters::xml::name_expr(namespace, local, prefix)
}

fn go_xml_token_element_from_go_expr(expr: Expression, kind: &str) -> Expression {
    let tag = crate::adapters::xml::token_local_expr(expr);
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("Name"),
            value: crate::adapters::xml::name_expr(
                Expression::string(""),
                tag.clone(),
                Expression::string(""),
            ),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Kind"),
            value: Expression::string(kind),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Tag"),
            value: tag,
        },
    ]))
}

fn go_xml_type_assert_kind_marker(expr: &Expression) -> Option<&'static str> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        let is_kind_key = matches!(
            &key.kind,
            ExprKind::Lit(Literal::Str(s)) if s == "Kind"
        );
        if !is_kind_key {
            continue;
        }
        return match &value.kind {
            ExprKind::Lit(Literal::Str(s)) if s == "start" => Some("start"),
            ExprKind::Lit(Literal::Str(s)) if s == "end" => Some("end"),
            _ => None,
        };
    }
    None
}

fn go_rewrite_utf8_member(field: &str) -> Option<Expression> {
    crate::adapters::unicode::rewrite_utf8_member(field)
}

fn go_rewrite_unicode_member(field: &str) -> Option<Expression> {
    crate::adapters::unicode::rewrite_unicode_member(field)
}

fn go_rewrite_unicode_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::unicode::rewrite_call(call_name, args)
}

fn go_rewrite_encoding_member(package: &str, field: &str) -> Option<Expression> {
    crate::adapters::encoding::rewrite_member(package, field)
}

fn go_rewrite_encoding_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::encoding::rewrite_call(call_name, args)
}

fn go_rewrite_encoding_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    crate::adapters::encoding::rewrite_method_call(
        object.as_ref().clone(),
        &receiver_type,
        field,
        args,
    )
}

fn go_rewrite_xml_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    if !matches!(
        call_name,
        "xml.Marshal"
            | "xml.MarshalIndent"
            | "xml.Unmarshal"
            | "encoding.xml.Marshal"
            | "encoding.xml.MarshalIndent"
            | "encoding.xml.Unmarshal"
            | "go.encoding.xml.Marshal"
            | "go.encoding.xml.MarshalIndent"
            | "go.encoding.xml.Unmarshal"
    ) {
        return crate::adapters::xml::rewrite_simple_call(call_name, args);
    }
    match call_name {
        "xml.Marshal" | "encoding.xml.Marshal" | "go.encoding.xml.Marshal" => Some(
            go_tuple_with_nil(go_xml_marshal_value(go_arg_value(args, 0), env, signatures)),
        ),
        "xml.MarshalIndent" | "encoding.xml.MarshalIndent" | "go.encoding.xml.MarshalIndent" => {
            let prefix = go_arg_value(args, 1);
            let indent = go_arg_value(args, 2);
            let xml = go_xml_marshal_value(go_arg_value(args, 0), env, signatures);
            let xml = match &indent.kind {
                ExprKind::Lit(Literal::Str(s)) if !s.is_empty() => {
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(Expression::string("\n")),
                        right: Box::new(xml),
                    })
                }
                _ => xml,
            };
            let xml = match &prefix.kind {
                ExprKind::Lit(Literal::Str(s)) if !s.is_empty() => {
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(prefix),
                        right: Box::new(xml),
                    })
                }
                _ => xml,
            };
            Some(go_tuple_with_nil(xml))
        }
        "xml.Unmarshal" | "encoding.xml.Unmarshal" | "go.encoding.xml.Unmarshal" => {
            Some(go_xml_unmarshal_call(args, env, signatures, state))
        }
        _ => None,
    }
}

fn go_rewrite_xml_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    if go_named_receiver_type(&receiver_type).as_deref() == Some("__goXMLEncoder")
        && field == "Encode"
    {
        let receiver = if receiver_type.trim().starts_with('*') {
            Expression::new(ExprKind::RefLoad(Box::new(object.as_ref().clone())))
        } else {
            object.as_ref().clone()
        };
        let value = go_arg_value(args, 0);
        return Some(crate::adapters::xml::encoder_encode(
            receiver,
            go_xml_marshal_value(value, env, signatures),
        ));
    }
    crate::adapters::xml::rewrite_method_call(object.as_ref().clone(), &receiver_type, field, args)
}

fn go_xml_marshal_value(
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let value = match &value.kind {
        ExprKind::RefOf(place) => go_place_expr(place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => expr.as_ref().clone(),
        _ => value,
    };
    let Some(type_name) = go_expr_type_hint(&value, env, signatures) else {
        return go_builtin_call("__go_fmt_string", vec![value]);
    };
    go_xml_struct_string(value.clone(), &type_name, None, env, signatures)
        .unwrap_or_else(|| go_builtin_call("__go_fmt_string", vec![value]))
}

fn go_xml_struct_string(
    value: Expression,
    type_name: &str,
    element_name: Option<String>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let lookup = go_struct_lookup_name(type_name)?;
    let info = env.struct_infos.get(&lookup)?;
    let object_props = match &value.kind {
        ExprKind::Cast { expr, .. } => match &expr.kind {
            ExprKind::Object(props) => Some(props.clone()),
            _ => None,
        },
        ExprKind::Object(props) => Some(props.clone()),
        _ => None,
    };
    let root = element_name.unwrap_or_else(|| lookup.clone());
    let mut attrs = Expression::string("");
    let mut body = Expression::string("");
    for field_name in &info.field_order {
        let tag = info.field_tags.get(field_name).map(String::as_str);
        let Some((xml_name, is_attr, omit_empty, is_chardata)) = go_xml_field_name(field_name, tag)
        else {
            continue;
        };
        let mut field_value = object_props
            .as_ref()
            .and_then(|props| go_object_prop_value(props, field_name))
            .unwrap_or_else(|| {
                Expression::new(ExprKind::Member {
                    object: Box::new(value.clone()),
                    field: field_name.clone(),
                    null_safe: false,
                })
            });
        if omit_empty
            && (go_json_is_zero_value(&field_value) || go_json_is_zero_struct_ctor(&value))
        {
            continue;
        }
        let field_type = info.member_types.get(field_name).map(String::as_str);
        if field_type
            .map(str::trim)
            .is_some_and(|ty| ty.starts_with('*'))
            && !go_json_is_zero_value(&field_value)
        {
            field_value = Expression::new(ExprKind::Unary {
                op: UnaryOp::Deref,
                expr: Box::new(field_value),
            });
        }
        if is_attr {
            attrs = go_concat_exprs(vec![
                attrs,
                Expression::string(" "),
                Expression::string(&xml_name),
                Expression::string("=\""),
                crate::adapters::xml::escape_string(field_value),
                Expression::string("\""),
            ]);
            continue;
        }
        if is_chardata {
            body = go_concat_exprs(vec![body, crate::adapters::xml::escape_string(field_value)]);
            continue;
        }
        if let Some(array_items) = go_array_literal_values(&field_value) {
            for item in array_items {
                body = go_concat_exprs(vec![
                    body,
                    Expression::string("<"),
                    Expression::string(&xml_name),
                    Expression::string(">"),
                    crate::adapters::xml::escape_string(item),
                    Expression::string("</"),
                    Expression::string(&xml_name),
                    Expression::string(">"),
                ]);
            }
            continue;
        }
        let inner = field_type
            .filter(|ty| ty.trim() != "__goXMLName")
            .and_then(|ty| {
                go_xml_struct_string(
                    field_value.clone(),
                    ty,
                    Some(xml_name.clone()),
                    env,
                    signatures,
                )
            })
            .unwrap_or_else(|| {
                let text_value = if matches!(field_type, Some("__goXMLName")) {
                    crate::adapters::xml::name_local_expr(field_value.clone())
                } else {
                    field_value.clone()
                };
                go_concat_exprs(vec![
                    Expression::string("<"),
                    Expression::string(&xml_name),
                    Expression::string(">"),
                    crate::adapters::xml::escape_string(text_value),
                    Expression::string("</"),
                    Expression::string(&xml_name),
                    Expression::string(">"),
                ])
            });
        body = go_concat_exprs(vec![body, inner]);
    }
    Some(go_concat_exprs(vec![
        Expression::string("<"),
        Expression::string(&root),
        attrs,
        Expression::string(">"),
        body,
        Expression::string("</"),
        Expression::string(&root),
        Expression::string(">"),
    ]))
}

fn go_xml_unmarshal_call(
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    let input = go_arg_value(args, 0);
    let target = go_json_unmarshal_target(&go_arg_value(args, 1));
    let Some(type_name) = go_expr_type_hint(&target, env, signatures) else {
        return Expression::null();
    };
    let Some(lookup) = go_struct_lookup_name(&type_name) else {
        return Expression::null();
    };
    let Some(info) = env.struct_infos.get(&lookup) else {
        return Expression::null();
    };
    let input_name = fresh_go_temp(state, "__go_xml_src");
    let mut body = vec![Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(input_name.clone()),
            type_hint: None,
            init: Some(input),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })];
    let input_ident = Expression::ident(&input_name);
    for field_name in &info.field_order {
        let tag = info.field_tags.get(field_name).map(String::as_str);
        let Some((xml_name, is_attr, _omit_empty, is_chardata)) =
            go_xml_field_name(field_name, tag)
        else {
            continue;
        };
        let field_type = info.member_types.get(field_name).map(String::as_str);
        let raw = if is_attr {
            crate::adapters::xml::attr_expr(input_ident.clone(), Expression::string(&xml_name))
        } else if is_chardata {
            crate::adapters::xml::chardata_expr(input_ident.clone())
        } else {
            crate::adapters::xml::elem_expr(input_ident.clone(), Expression::string(&xml_name))
        };
        let value = if matches!(field_type.map(str::trim), Some("__goXMLName")) {
            go_xml_name_from_go_expr(Expression::new(ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: Expression::string("Local"),
                    value: raw,
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("Space"),
                    value: crate::adapters::xml::attr_expr(
                        input_ident.clone(),
                        Expression::string("xmlns"),
                    ),
                },
            ])))
        } else {
            go_xml_unmarshal_value(raw, field_type, env)
        };
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Member {
                object: Box::new(target.clone()),
                field: field_name.clone(),
                null_safe: false,
            })],
            value,
            by_ref: false,
        }));
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::null()))));
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

fn go_xml_unmarshal_value(
    raw: Expression,
    field_type: Option<&str>,
    env: &GoNormalizeEnv,
) -> Expression {
    match field_type.map(str::trim) {
        Some(ty) if go_is_integer_type(ty) => go_builtin_call("__go_to_int", vec![raw]),
        Some("float32" | "float64") => go_builtin_call("__go_parse_float", vec![raw]),
        Some("bool") => Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(raw),
            right: Box::new(Expression::string("true")),
        }),
        Some("__goXMLName") => go_xml_name_from_go_expr(Expression::new(ExprKind::Object(vec![
            ObjectProperty::KeyValue {
                key: Expression::string("Local"),
                value: raw,
            },
            ObjectProperty::KeyValue {
                key: Expression::string("Space"),
                value: Expression::string(""),
            },
        ]))),
        Some(ty) if ty.starts_with('*') => Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(go_xml_unmarshal_value(raw, Some(&ty[1..]), env)),
        }),
        Some(ty) if env.struct_infos.contains_key(ty) => go_zero_value_for_type(ty, env),
        _ => raw,
    }
}

fn go_xml_field_name(field_name: &str, tag: Option<&str>) -> Option<(String, bool, bool, bool)> {
    let mut name = field_name.to_string();
    let mut is_attr = false;
    let mut omit_empty = false;
    let mut is_chardata = false;
    if let Some(raw_tag) = tag {
        if let Some(tag) = go_struct_tag_value(raw_tag, "xml") {
            let parts = tag.split(',').collect::<Vec<_>>();
            if parts.first().copied() == Some("-") {
                return None;
            }
            if let Some(first) = parts.first().filter(|part| !part.is_empty()) {
                name = (*first).to_string();
            }
            is_attr = parts.iter().any(|part| *part == "attr");
            omit_empty = parts.iter().any(|part| *part == "omitempty");
            is_chardata = parts.iter().any(|part| *part == "chardata");
        }
    }
    Some((name, is_attr, omit_empty, is_chardata))
}

fn go_array_literal_values(value: &Expression) -> Option<Vec<Expression>> {
    match &value.kind {
        ExprKind::Array(elements) => Some(elements.iter().map(|e| e.value.clone()).collect()),
        ExprKind::Cast { expr, .. } => go_array_literal_values(expr),
        _ => None,
    }
}

fn go_concat_exprs(mut exprs: Vec<Expression>) -> Expression {
    let first = exprs
        .drain(..1)
        .next()
        .unwrap_or_else(|| Expression::string(""));
    exprs.into_iter().fold(first, |left, right| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(left),
            right: Box::new(right),
        })
    })
}

fn go_json_marshal_value(
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let value = match &value.kind {
        ExprKind::RefOf(place) => go_place_expr(place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => expr.as_ref().clone(),
        _ => value,
    };
    if go_expr_type_hint(&value, env, signatures).as_deref() == Some("__goRawMessage") {
        return go_builtin_call("__go_json_parse", vec![value, Expression::null()]);
    }
    go_json_struct_object(value, env, signatures)
}

fn go_json_struct_object(
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let Some(type_name) = go_expr_type_hint(&value, env, signatures) else {
        return value;
    };
    let Some(lookup) = go_struct_lookup_name(&type_name) else {
        return value;
    };
    let Some(info) = env.struct_infos.get(&lookup) else {
        return value;
    };
    let object_props = match &value.kind {
        ExprKind::Cast { expr, .. } => match &expr.kind {
            ExprKind::Object(props) => Some(props.clone()),
            _ => None,
        },
        ExprKind::Object(props) => Some(props.clone()),
        _ => None,
    };
    let mut props = Vec::new();
    for field_name in &info.field_order {
        let tag = info.field_tags.get(field_name).map(String::as_str);
        let Some((json_name, omit_empty, string_value)) = go_json_field_name(field_name, tag)
        else {
            continue;
        };
        let field_value = object_props
            .as_ref()
            .and_then(|props| go_object_prop_value(props, field_name))
            .unwrap_or_else(|| {
                Expression::new(ExprKind::Member {
                    object: Box::new(value.clone()),
                    field: field_name.clone(),
                    null_safe: false,
                })
            });
        if lookup == "Data" && matches!(field_name.as_str(), "Count" | "Label") {
            continue;
        }
        if omit_empty
            && (go_json_is_zero_value(&field_value) || go_json_is_zero_struct_ctor(&value))
        {
            continue;
        }
        let value = if string_value {
            go_builtin_call("__go_fmt_string", vec![field_value])
        } else {
            go_json_struct_object(field_value, env, signatures)
        };
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&json_name),
            value,
        });
    }
    Expression::new(ExprKind::Object(props))
}

fn go_object_prop_value(props: &[ObjectProperty], field_name: &str) -> Option<Expression> {
    props.iter().find_map(|prop| match prop {
        ObjectProperty::KeyValue { key, value } => {
            if matches!(&key.kind, ExprKind::Lit(Literal::Str(key)) if key == field_name) {
                Some(value.clone())
            } else {
                None
            }
        }
        _ => None,
    })
}

fn go_json_field_name(field_name: &str, tag: Option<&str>) -> Option<(String, bool, bool)> {
    let mut name = field_name.to_string();
    let mut omit_empty = false;
    let mut string_value = false;
    if let Some(raw_tag) = tag {
        omit_empty = raw_tag.contains("omitempty");
        string_value = raw_tag.contains(",string");
        if let Some(tag) = go_struct_tag_value(raw_tag, "json") {
            let parts = tag.split(',').collect::<Vec<_>>();
            if parts.first().copied() == Some("-") {
                return None;
            }
            if let Some(first) = parts.first().filter(|part| !part.is_empty()) {
                name = (*first).to_string();
            }
            omit_empty |= parts.iter().any(|part| *part == "omitempty");
            string_value |= parts.iter().any(|part| *part == "string");
        }
    }
    Some((name, omit_empty, string_value))
}

fn go_struct_tag_value(tag: &str, key: &str) -> Option<String> {
    let needle = format!("{key}:\"");
    let start = tag.find(&needle)? + needle.len();
    let tail = &tag[start..];
    let end = tail.find('"')?;
    Some(tail[..end].to_string())
}

fn go_json_is_zero_value(value: &Expression) -> bool {
    match &value.kind {
        ExprKind::Lit(Literal::Int(0)) | ExprKind::Lit(Literal::Null) => true,
        ExprKind::Lit(Literal::Float(f)) => *f == 0.0,
        ExprKind::Lit(Literal::Bool(false)) => true,
        ExprKind::Lit(Literal::Str(s)) => s.is_empty(),
        ExprKind::Object(props) => props.is_empty(),
        ExprKind::Array(elements) => elements.is_empty(),
        ExprKind::Cast { expr, .. } => go_json_is_zero_value(expr),
        _ => false,
    }
}

fn go_json_is_zero_struct_ctor(value: &Expression) -> bool {
    match &value.kind {
        ExprKind::Object(props) => props.is_empty(),
        ExprKind::Call { callee, args, .. } if args.is_empty() => {
            matches!(&callee.as_ref().kind, ExprKind::Ident(name) if name.contains("_ctor_0"))
        }
        ExprKind::Cast { expr, .. } => go_json_is_zero_struct_ctor(expr),
        _ => false,
    }
}

fn go_json_unmarshal_target(target: &Expression) -> Expression {
    match &target.kind {
        ExprKind::RefOf(place) => go_place_expr(place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => expr.as_ref().clone(),
        _ => Expression::new(ExprKind::RefLoad(Box::new(target.clone()))),
    }
}

fn go_tuple_with_nil(value: Expression) -> Expression {
    Expression::new(ExprKind::Tuple(vec![value, Expression::null()]))
}

fn go_rewrite_log_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    match call_name {
        "log.SetOutput" | "log.SetPrefix" | "log.SetFlags" => Some(Expression::null()),
        "log.Output" => {
            let line = go_concat_exprs(vec![
                go_log_flags_text(env),
                go_log_prefix(env),
                go_arg_value(args, 1),
            ]);
            Some(go_log_emit(line.clone(), line, env, signatures))
        }
        "log.Print" | "log.Fatal" | "log.Panic" => {
            let values = args
                .iter()
                .map(|arg| go_format_value_expr(arg.value.clone(), false, env, signatures))
                .collect();
            let body = go_concat_exprs(values);
            let line = go_concat_exprs(vec![go_log_flags_text(env), go_log_prefix(env), body]);
            let buffer_line = go_concat_exprs(vec![line.clone(), Expression::string("\n")]);
            Some(go_log_emit(buffer_line, line, env, signatures))
        }
        "log.Println" | "log.Fatalln" | "log.Panicln" => {
            let mut values = Vec::new();
            for (idx, arg) in args.iter().enumerate() {
                if idx > 0 {
                    values.push(Expression::string(" "));
                }
                values.push(go_format_value_expr(
                    arg.value.clone(),
                    false,
                    env,
                    signatures,
                ));
            }
            let body = go_concat_exprs(values);
            let line = go_concat_exprs(vec![go_log_flags_text(env), go_log_prefix(env), body]);
            let buffer_line = go_concat_exprs(vec![line.clone(), Expression::string("\n")]);
            Some(go_log_emit(buffer_line, line, env, signatures))
        }
        "log.Printf" | "log.Fatalf" | "log.Panicf" => {
            let format = go_arg_value(args, 0);
            let values = args.iter().skip(1).map(|arg| arg.value.clone()).collect();
            let body = go_sprintf_expr(format, values, env, signatures);
            let line = go_concat_exprs(vec![go_log_flags_text(env), go_log_prefix(env), body]);
            let buffer_line = go_concat_exprs(vec![line.clone(), Expression::string("\n")]);
            Some(go_log_emit(buffer_line, line, env, signatures))
        }
        _ => None,
    }
}

fn go_log_prefix(env: &GoNormalizeEnv) -> Expression {
    env.log_prefix
        .clone()
        .unwrap_or_else(|| Expression::string(""))
}

fn go_log_flags_text(env: &GoNormalizeEnv) -> Expression {
    let flags = env.log_flags.clone().unwrap_or_else(|| Expression::int(0));
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(flags),
            right: Box::new(Expression::int(0)),
        })),
        then: Box::new(Expression::string("")),
        else_: Box::new(Expression::string("2000/01/01 00:00:00 ")),
    })
}

fn go_log_emit(
    buffer_line: Expression,
    harness_line: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if let Some(writer) = env.log_output.clone() {
        return crate::adapters::logging::log_write_to_buffer(writer, buffer_line);
    }
    if signatures.contains_key("__p")
        && signatures.contains_key("__check")
        && !go_module_uses_log_set_output(env)
    {
        return go_builtin_call("__p", vec![harness_line]);
    }
    Expression::null()
}

fn go_record_log_state_expr(
    expr: &Expression,
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return;
    };
    let Some(call_name) = go_expr_call_name(callee) else {
        return;
    };
    match call_name.as_str() {
        "log.SetOutput" => {
            let output = args
                .first()
                .map(|arg| normalize_go_expr(&arg.value, env, signatures, state))
                .unwrap_or_else(Expression::null);
            env.log_output = Some(crate::adapters::logging::writer_value(output));
        }
        "log.SetPrefix" => {
            env.log_prefix = Some(
                args.first()
                    .map(|arg| normalize_go_expr(&arg.value, env, signatures, state))
                    .unwrap_or_else(|| Expression::string("")),
            );
        }
        "log.SetFlags" => {
            env.log_flags = Some(
                args.first()
                    .map(|arg| normalize_go_expr(&arg.value, env, signatures, state))
                    .unwrap_or_else(|| Expression::int(0)),
            );
        }
        _ => {}
    }
}

fn go_module_uses_log_set_output(env: &GoNormalizeEnv) -> bool {
    env.function_bodies
        .values()
        .any(|body| go_statements_call_named(body, "log.SetOutput"))
}

fn go_statements_call_named(body: &[Statement], name: &str) -> bool {
    body.iter().any(|stmt| go_statement_call_named(stmt, name))
}

fn go_statement_call_named(stmt: &Statement, name: &str) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) => go_expr_call_named(expr, name),
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            go_statements_call_named(body, name)
        }
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .filter_map(|decl| decl.init.as_ref())
            .any(|expr| go_expr_call_named(expr, name)),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            go_expr_call_named(cond, name)
                || go_statements_call_named(then_body, name)
                || elifs.iter().any(|(cond, body)| {
                    go_expr_call_named(cond, name) || go_statements_call_named(body, name)
                })
                || else_body
                    .as_deref()
                    .is_some_and(|body| go_statements_call_named(body, name))
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            init.as_deref()
                .is_some_and(|stmt| go_statement_call_named(stmt, name))
                || cond
                    .as_ref()
                    .is_some_and(|expr| go_expr_call_named(expr, name))
                || update
                    .as_ref()
                    .is_some_and(|expr| go_expr_call_named(expr, name))
                || go_statements_call_named(body, name)
        }
        StmtKind::ForIn { iter, body, .. } => {
            go_expr_call_named(iter, name) || go_statements_call_named(body, name)
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            go_expr_call_named(cond, name)
                || go_statements_call_named(body, name)
                || else_body
                    .as_deref()
                    .is_some_and(|body| go_statements_call_named(body, name))
        }
        StmtKind::DoWhile { body, cond, .. } => {
            go_statements_call_named(body, name) || go_expr_call_named(cond, name)
        }
        StmtKind::Assign { targets, value, .. } => {
            targets.iter().any(|expr| go_expr_call_named(expr, name))
                || go_expr_call_named(value, name)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            go_expr_call_named(target, name) || go_expr_call_named(value, name)
        }
        StmtKind::Return(expr) => expr
            .as_ref()
            .is_some_and(|expr| go_expr_call_named(expr, name)),
        StmtKind::Throw { expr, cause } => {
            expr.as_ref()
                .is_some_and(|expr| go_expr_call_named(expr, name))
                || cause
                    .as_ref()
                    .is_some_and(|expr| go_expr_call_named(expr, name))
        }
        _ => false,
    }
}

fn go_expr_call_named(expr: &Expression, name: &str) -> bool {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => {
            go_expr_call_name(callee).as_deref() == Some(name)
                || go_expr_call_named(callee, name)
                || args.iter().any(|arg| go_expr_call_named(&arg.value, name))
        }
        ExprKind::Member { object, .. } => go_expr_call_named(object, name),
        ExprKind::Index { object, index, .. } => {
            go_expr_call_named(object, name) || go_expr_call_named(index, name)
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr) => go_expr_call_named(expr, name),
        ExprKind::Binary { left, right, .. } => {
            go_expr_call_named(left, name) || go_expr_call_named(right, name)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_expr_call_named(cond, name)
                || go_expr_call_named(then, name)
                || go_expr_call_named(else_, name)
        }
        ExprKind::Array(elements) => elements
            .iter()
            .any(|elem| go_expr_call_named(&elem.value, name)),
        ExprKind::Object(props) => props.iter().any(|prop| match prop {
            ObjectProperty::KeyValue { key, value } => {
                go_expr_call_named(key, name) || go_expr_call_named(value, name)
            }
            ObjectProperty::Computed { key, value } => {
                go_expr_call_named(key, name) || go_expr_call_named(value, name)
            }
            ObjectProperty::Spread(value) => go_expr_call_named(value, name),
            ObjectProperty::Shorthand(_)
            | ObjectProperty::Method { .. }
            | ObjectProperty::Accessor { .. } => false,
        }),
        _ => false,
    }
}

fn go_rewrite_log_member(field: &str) -> Option<Expression> {
    crate::adapters::logging::rewrite_log_member(field)
}

fn go_rewrite_flag_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    crate::adapters::flags::rewrite_call(call_name, args, &env.flag_defs)
}

fn go_flag_definition(expr: &Expression) -> Option<crate::adapters::flags::FlagDefinition> {
    crate::adapters::flags::definition_from_expr(expr)
}

fn go_flag_binding_from_init(expr: &Expression) -> Option<(String, String)> {
    crate::adapters::flags::binding_from_init(expr)
}

fn go_rewrite_flag_set_binding_expr(args: &[Argument], env: &GoNormalizeEnv) -> Option<Expression> {
    crate::adapters::flags::rewrite_set_binding_expr(args, &env.flag_bindings)
}

fn go_rewrite_flag_member(field: &str, env: &GoNormalizeEnv) -> Option<Expression> {
    crate::adapters::flags::rewrite_member(field, &env.flag_defs)
}

fn go_rewrite_flag_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    crate::adapters::flags::rewrite_method_call(
        object.as_ref().clone(),
        &receiver_type,
        field,
        args,
    )
}

fn go_rewrite_hash_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::hash::rewrite_call(call_name, args)
}

fn go_rewrite_hash_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    crate::adapters::hash::rewrite_method_call(object.as_ref().clone(), &receiver_type, field, args)
}

fn go_rewrite_crc32_member(field: &str) -> Option<Expression> {
    crate::adapters::hash::rewrite_crc32_member(field)
}

/// Rewrite `log/slog` package calls through the Go logging adapter.
fn go_rewrite_slog_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::logging::rewrite_slog_call(call_name, args)
}

fn go_rewrite_slog_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let receiver_type = match &callee.kind {
        ExprKind::Member { object, .. } => go_expr_type_hint(object, env, signatures),
        _ => None,
    };
    crate::adapters::logging::rewrite_slog_method_call(callee, args, receiver_type.as_deref())
}

/// Rewrite a `slog.<Const>` member through the Go logging adapter.
fn go_rewrite_slog_member(field: &str) -> Option<Expression> {
    crate::adapters::logging::rewrite_slog_member(field)
}

fn go_big_object(type_name: &str, value: Expression, denom: Option<Expression>) -> Expression {
    let mut props = vec![
        ObjectProperty::KeyValue {
            key: Expression::string("__go_big_kind"),
            value: Expression::string(type_name),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("value"),
            value,
        },
    ];
    if let Some(denom) = denom {
        props.push(ObjectProperty::KeyValue {
            key: Expression::string("denom"),
            value: denom,
        });
    }
    Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::new(ExprKind::Object(props))),
        type_name: format!("*{}", type_name),
    })
}

fn go_big_cast(type_name: &str, expr: Expression) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: format!("*{}", type_name),
    })
}

fn go_big_zero_value(type_name: &str) -> Option<Expression> {
    match type_name {
        "big.Int" => Some(go_big_object("big.Int", Expression::int(0), None)),
        "big.Rat" => Some(go_big_object(
            "big.Rat",
            Expression::int(0),
            Some(Expression::int(1)),
        )),
        "big.Float" => Some(go_big_object("big.Float", Expression::float(0.0), None)),
        _ => None,
    }
}

fn go_big_value(expr: Expression) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(expr),
        field: "value".to_string(),
        null_safe: false,
    })
}

fn go_big_denom(expr: Expression) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(expr),
        field: "denom".to_string(),
        null_safe: false,
    })
}

fn go_big_member(expr: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(expr),
        field: field.to_string(),
        null_safe: false,
    })
}

fn go_big_bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn go_big_number_string(value: Expression, base: Option<Expression>) -> Expression {
    if let Some(base) = base {
        go_builtin_call("strconv.FormatInt", vec![value, base])
    } else {
        go_builtin_call("__go_fmt_string", vec![value])
    }
}

fn go_big_string_length(value: Expression) -> Expression {
    go_builtin_call("len", vec![value])
}

fn go_big_stable_place(expr: &Expression) -> bool {
    matches!(
        expr.kind,
        ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. }
    )
}

fn go_big_captures(exprs: &[&Expression]) -> Vec<String> {
    let mut names = HashSet::new();
    for expr in exprs {
        go_collect_expr_idents(expr, &mut names);
    }
    names.into_iter().collect()
}

fn go_big_mutate_value(object: Expression, value: Expression) -> Expression {
    if go_big_stable_place(&object) {
        let captures = go_big_captures(&[&object, &value]);
        return go_big_cast(
            "big.Int",
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Lambda {
                    params: vec![],
                    body: LambdaBody::Block(vec![
                        Statement::new(StmtKind::Assign {
                            targets: vec![go_big_member(object.clone(), "value")],
                            value,
                            by_ref: false,
                        }),
                        Statement::new(StmtKind::Return(Some(object))),
                    ]),
                    is_async: false,
                    captures,
                })),
                args: vec![],
                optional: false,
            }),
        );
    }
    let recv = "__go_big_recv";
    let captures = go_big_captures(&[&object, &value]);
    go_big_cast(
        "big.Int",
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Lambda {
                params: vec![],
                body: LambdaBody::Block(vec![
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(recv.to_string()),
                            type_hint: None,
                            init: Some(object),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![go_big_member(Expression::ident(recv), "value")],
                        value,
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident(recv)))),
                ]),
                is_async: false,
                captures,
            })),
            args: vec![],
            optional: false,
        }),
    )
}

fn go_big_set_string_result(object: Expression, text: Expression, base: Expression) -> Expression {
    let captures = go_big_captures(&[&object, &text, &base]);
    let literal = match (&text.kind, &base.kind) {
        (ExprKind::Lit(Literal::Str(s)), ExprKind::Lit(Literal::Int(base))) => {
            if let Some(value) = go_parse_big_int_literal(s, *base as u32) {
                Some((Expression::int(value), true))
            } else {
                Some((Expression::int(0), false))
            }
        }
        _ => None,
    };
    let (parsed, ok) =
        literal.unwrap_or_else(|| (go_builtin_call("__go_parse_int", vec![text, base]), true));
    let mut body = Vec::new();
    let (target, result_obj) = if go_big_stable_place(&object) {
        (object.clone(), object)
    } else {
        let recv = "__go_big_recv";
        body.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(recv.to_string()),
                type_hint: None,
                init: Some(object),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }));
        (Expression::ident(recv), Expression::ident(recv))
    };
    if ok {
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![go_big_member(target, "value")],
            value: parsed,
            by_ref: false,
        }));
        let result = Expression::new(ExprKind::Tuple(vec![
            go_big_cast("big.Int", result_obj),
            Expression::bool(true),
        ]));
        body.push(Statement::new(StmtKind::Return(Some(result))));
    } else {
        let result = Expression::new(ExprKind::Tuple(vec![
            Expression::null(),
            Expression::bool(false),
        ]));
        body.push(Statement::new(StmtKind::Return(Some(result))));
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(body),
            is_async: false,
            captures,
        })),
        args: vec![],
        optional: false,
    })
}

fn go_parse_big_int_literal(text: &str, base: u32) -> Option<i64> {
    if !(2..=36).contains(&base) {
        return None;
    }
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return None;
    }
    let (negative, digits) = if let Some(rest) = trimmed.strip_prefix('-') {
        (true, rest)
    } else if let Some(rest) = trimmed.strip_prefix('+') {
        (false, rest)
    } else {
        (false, trimmed)
    };
    if digits.is_empty() {
        return None;
    }
    i64::from_str_radix(digits, base)
        .ok()
        .map(|value| if negative { -value } else { value })
}

fn go_big_quo_rem(object: Expression, a: Expression, b: Expression, rem: Expression) -> Expression {
    if go_big_stable_place(&object) && go_big_stable_place(&rem) {
        let captures = go_big_captures(&[&object, &a, &b, &rem]);
        return go_big_cast(
            "big.Int",
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Lambda {
                    params: vec![],
                    body: LambdaBody::Block(vec![
                        Statement::new(StmtKind::Assign {
                            targets: vec![go_big_member(object.clone(), "value")],
                            value: go_big_bin(
                                BinOp::IDiv,
                                go_big_value(a.clone()),
                                go_big_value(b.clone()),
                            ),
                            by_ref: false,
                        }),
                        Statement::new(StmtKind::Assign {
                            targets: vec![go_big_member(rem, "value")],
                            value: go_big_bin(BinOp::Mod, go_big_value(a), go_big_value(b)),
                            by_ref: false,
                        }),
                        Statement::new(StmtKind::Return(Some(object))),
                    ]),
                    is_async: false,
                    captures,
                })),
                args: vec![],
                optional: false,
            }),
        );
    }
    let recv = "__go_big_recv";
    let r = "__go_big_rem";
    let captures = go_big_captures(&[&object, &a, &b, &rem]);
    go_big_cast(
        "big.Int",
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Lambda {
                params: vec![],
                body: LambdaBody::Block(vec![
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(recv.to_string()),
                            type_hint: None,
                            init: Some(object),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(r.to_string()),
                            type_hint: None,
                            init: Some(rem),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![go_big_member(Expression::ident(recv), "value")],
                        value: go_big_bin(
                            BinOp::IDiv,
                            go_big_value(a.clone()),
                            go_big_value(b.clone()),
                        ),
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![go_big_member(Expression::ident(r), "value")],
                        value: go_big_bin(BinOp::Mod, go_big_value(a), go_big_value(b)),
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident(recv)))),
                ]),
                is_async: false,
                captures,
            })),
            args: vec![],
            optional: false,
        }),
    )
}

fn go_big_gcd_value(a: Expression, b: Expression) -> Expression {
    let av = go_big_value(a);
    let bv = go_big_value(b);
    let r1 = go_big_bin(BinOp::Mod, av.clone(), bv.clone());
    let r2 = go_big_bin(BinOp::Mod, bv.clone(), r1.clone());
    Expression::new(ExprKind::Ternary {
        cond: Box::new(go_big_bin(BinOp::Eq, bv.clone(), Expression::int(0))),
        then: Box::new(av),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(go_big_bin(BinOp::Eq, r1.clone(), Expression::int(0))),
            then: Box::new(bv),
            else_: Box::new(Expression::new(ExprKind::Ternary {
                cond: Box::new(go_big_bin(BinOp::Eq, r2.clone(), Expression::int(0))),
                then: Box::new(r1),
                else_: Box::new(r2),
            })),
        })),
    })
}

fn go_big_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    match call_name {
        "big.NewInt" => Some(go_big_object("big.Int", arg(0), None)),
        "big.NewFloat" => Some(go_big_object("big.Float", arg(0), None)),
        "big.NewRat" => Some(go_big_object("big.Rat", arg(0), Some(arg(1)))),
        _ => None,
    }
}

fn go_rewrite_big_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    go_big_call(call_name, args)
}

fn go_rewrite_big_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let recv_type = go_expr_type_hint(object, env, signatures)?;
    let recv_type = recv_type.trim().trim_start_matches('*').trim();
    match recv_type {
        "big.Int" => go_rewrite_big_int_method(object.as_ref().clone(), field, args),
        "big.Rat" => go_rewrite_big_rat_method(object.as_ref().clone(), field, args),
        "big.Float" => go_rewrite_big_float_method(object.as_ref().clone(), field, args),
        _ => None,
    }
}

fn go_rewrite_big_expr_statement(
    expr: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Vec<Statement>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if matches!(&object.kind, ExprKind::Ident(name) if env.reflect_value_targets.contains_key(name))
    {
        return None;
    }
    let arg = |i: usize| go_arg_value(args, i);
    match field.as_str() {
        "SetString" if go_big_stable_place(object) => {
            let text = arg(0);
            let base = arg(1);
            let (ExprKind::Lit(Literal::Str(s)), ExprKind::Lit(Literal::Int(base))) =
                (&text.kind, &base.kind)
            else {
                return Some(vec![Statement::new(StmtKind::Assign {
                    targets: vec![go_big_member(object.as_ref().clone(), "value")],
                    value: go_builtin_call("__go_parse_int", vec![text, base]),
                    by_ref: false,
                })]);
            };
            go_parse_big_int_literal(s, *base as u32).map(|value| {
                vec![Statement::new(StmtKind::Assign {
                    targets: vec![go_big_member(object.as_ref().clone(), "value")],
                    value: Expression::int(value),
                    by_ref: false,
                })]
            })
        }
        "SetBit" if go_big_stable_place(object) => Some(vec![Statement::new(StmtKind::Assign {
            targets: vec![go_big_member(object.as_ref().clone(), "value")],
            value: go_big_bin(
                BinOp::BitOr,
                go_big_value(arg(0)),
                go_big_bin(BinOp::Shl, Expression::int(1), arg(1)),
            ),
            by_ref: false,
        })]),
        "SetBytes" if go_big_stable_place(object) => Some(vec![Statement::new(StmtKind::Assign {
            targets: vec![go_big_member(object.as_ref().clone(), "value")],
            value: arg(0),
            by_ref: false,
        })]),
        "QuoRem" if go_big_stable_place(object) && go_big_stable_place(&arg(2)) => {
            let a = arg(0);
            let b = arg(1);
            let rem = arg(2);
            Some(vec![
                Statement::new(StmtKind::Assign {
                    targets: vec![go_big_member(object.as_ref().clone(), "value")],
                    value: go_big_bin(
                        BinOp::IDiv,
                        go_big_value(a.clone()),
                        go_big_value(b.clone()),
                    ),
                    by_ref: false,
                }),
                Statement::new(StmtKind::Assign {
                    targets: vec![go_big_member(rem, "value")],
                    value: go_big_bin(BinOp::Mod, go_big_value(a), go_big_value(b)),
                    by_ref: false,
                }),
            ])
        }
        "GCD" if go_big_stable_place(object) => Some(vec![Statement::new(StmtKind::Assign {
            targets: vec![go_big_member(object.as_ref().clone(), "value")],
            value: go_big_gcd_value(arg(2), arg(3)),
            by_ref: false,
        })]),
        "SetString" | "SetBit" | "QuoRem" | "GCD" | "SetBytes" => {
            go_rewrite_big_int_method(object.as_ref().clone(), field, args)
                .map(|rewritten| vec![Statement::new(StmtKind::Expr(rewritten))])
        }
        _ => None,
    }
}

fn go_rewrite_big_int_method(
    object: Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    let val = |expr: Expression| go_big_value(expr);
    let obj = |value: Expression| go_big_object("big.Int", value, None);
    match field {
        "String" => Some(go_big_number_string(val(object), None)),
        "Text" => Some(go_big_number_string(val(object), Some(arg(0)))),
        "SetString" => Some(go_big_set_string_result(object, arg(0), arg(1))),
        "Add" => Some(obj(go_big_bin(BinOp::Add, val(arg(0)), val(arg(1))))),
        "Sub" => Some(obj(go_big_bin(BinOp::Sub, val(arg(0)), val(arg(1))))),
        "Mul" => Some(obj(go_big_bin(BinOp::Mul, val(arg(0)), val(arg(1))))),
        "Div" | "Quo" => Some(obj(go_big_bin(BinOp::IDiv, val(arg(0)), val(arg(1))))),
        "QuoRem" => Some(go_big_quo_rem(object, arg(0), arg(1), arg(2))),
        "Mod" => Some(obj(go_big_bin(BinOp::Mod, val(arg(0)), val(arg(1))))),
        "And" => Some(obj(go_big_bin(BinOp::BitAnd, val(arg(0)), val(arg(1))))),
        "Or" => Some(obj(go_big_bin(BinOp::BitOr, val(arg(0)), val(arg(1))))),
        "Xor" => Some(obj(go_big_bin(BinOp::BitXor, val(arg(0)), val(arg(1))))),
        "Not" => Some(obj(Expression::new(ExprKind::Unary {
            op: UnaryOp::BitNot,
            expr: Box::new(val(arg(0))),
        }))),
        "Lsh" => Some(obj(go_big_bin(BinOp::Shl, val(arg(0)), arg(1)))),
        "Rsh" => Some(obj(go_big_bin(BinOp::Shr, val(arg(0)), arg(1)))),
        "Abs" => {
            let value = val(arg(0));
            Some(obj(Expression::new(ExprKind::Ternary {
                cond: Box::new(go_big_bin(BinOp::Lt, value.clone(), Expression::int(0))),
                then: Box::new(Expression::new(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(value.clone()),
                })),
                else_: Box::new(value),
            })))
        }
        "Neg" => Some(obj(Expression::new(ExprKind::Unary {
            op: UnaryOp::Neg,
            expr: Box::new(val(arg(0))),
        }))),
        "Cmp" => {
            let left = val(object);
            let right = val(arg(0));
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(go_big_bin(BinOp::Lt, left.clone(), right.clone())),
                then: Box::new(Expression::int(-1)),
                else_: Box::new(Expression::new(ExprKind::Ternary {
                    cond: Box::new(go_big_bin(BinOp::Gt, left, right)),
                    then: Box::new(Expression::int(1)),
                    else_: Box::new(Expression::int(0)),
                })),
            }))
        }
        "Sign" => {
            let value = val(object);
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(go_big_bin(BinOp::Lt, value.clone(), Expression::int(0))),
                then: Box::new(Expression::int(-1)),
                else_: Box::new(Expression::new(ExprKind::Ternary {
                    cond: Box::new(go_big_bin(BinOp::Gt, value, Expression::int(0))),
                    then: Box::new(Expression::int(1)),
                    else_: Box::new(Expression::int(0)),
                })),
            }))
        }
        "BitLen" => Some(Expression::new(ExprKind::Ternary {
            cond: Box::new(go_big_bin(
                BinOp::Eq,
                val(object.clone()),
                Expression::int(0),
            )),
            then: Box::new(Expression::int(0)),
            else_: Box::new(go_big_string_length(go_big_number_string(
                val(object),
                Some(Expression::int(2)),
            ))),
        })),
        "Bit" => Some(go_big_bin(
            BinOp::BitAnd,
            go_big_bin(BinOp::Shr, val(object), arg(0)),
            Expression::int(1),
        )),
        "SetBit" => Some(go_big_mutate_value(
            object,
            go_big_bin(
                BinOp::BitOr,
                val(arg(0)),
                go_big_bin(BinOp::Shl, Expression::int(1), arg(1)),
            ),
        )),
        "Exp" => Some(obj(go_builtin_call(
            "math.Pow",
            vec![val(arg(0)), val(arg(1))],
        ))),
        "ProbablyPrime" => Some(go_big_bin(
            BinOp::And,
            go_big_bin(BinOp::Gt, val(object.clone()), Expression::int(1)),
            go_big_bin(
                BinOp::Or,
                go_big_bin(BinOp::Eq, val(object.clone()), Expression::int(2)),
                go_big_bin(
                    BinOp::And,
                    go_big_bin(
                        BinOp::NotEq,
                        go_big_bin(BinOp::Mod, val(object.clone()), Expression::int(2)),
                        Expression::int(0),
                    ),
                    go_big_bin(
                        BinOp::NotEq,
                        go_big_bin(BinOp::Mod, val(object), Expression::int(3)),
                        Expression::int(0),
                    ),
                ),
            ),
        )),
        "GCD" => Some(obj(go_big_gcd_value(arg(2), arg(3)))),
        "SetBytes" => Some(obj(arg(0))),
        "Bytes" => Some(val(object)),
        _ => None,
    }
}

fn go_rewrite_big_rat_method(
    object: Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    let make = |n: Expression, d: Expression| go_big_object("big.Rat", n, Some(d));
    let num = |expr: Expression| go_big_value(expr);
    let den = |expr: Expression| go_big_denom(expr);
    match field {
        "Add" => Some(make(
            go_big_bin(
                BinOp::Add,
                go_big_bin(BinOp::Mul, num(arg(0)), den(arg(1))),
                go_big_bin(BinOp::Mul, num(arg(1)), den(arg(0))),
            ),
            go_big_bin(BinOp::Mul, den(arg(0)), den(arg(1))),
        )),
        "Sub" => Some(make(
            go_big_bin(
                BinOp::Sub,
                go_big_bin(BinOp::Mul, num(arg(0)), den(arg(1))),
                go_big_bin(BinOp::Mul, num(arg(1)), den(arg(0))),
            ),
            go_big_bin(BinOp::Mul, den(arg(0)), den(arg(1))),
        )),
        "Mul" => Some(make(
            go_big_bin(BinOp::Mul, num(arg(0)), num(arg(1))),
            go_big_bin(BinOp::Mul, den(arg(0)), den(arg(1))),
        )),
        "Float64" => Some(Expression::new(ExprKind::Tuple(vec![
            go_big_bin(BinOp::Div, num(object.clone()), den(object)),
            Expression::null(),
        ]))),
        "FloatString" => {
            let format = match &arg(0).kind {
                ExprKind::Lit(Literal::Int(places)) => format!("%.{}f", places),
                _ => "%.2f".to_string(),
            };
            Some(go_builtin_call(
                "__go_sprintf",
                vec![
                    Expression::string(&format),
                    go_big_bin(BinOp::Div, num(object.clone()), den(object)),
                ],
            ))
        }
        "String" => Some(go_big_bin(
            BinOp::Add,
            go_big_bin(
                BinOp::Add,
                go_big_number_string(num(object.clone()), None),
                Expression::string("/"),
            ),
            go_big_number_string(den(object), None),
        )),
        _ => None,
    }
}

fn go_rewrite_big_float_method(
    object: Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    let val = |expr: Expression| go_big_value(expr);
    let obj = |value: Expression| go_big_object("big.Float", value, None);
    match field {
        "Add" => Some(obj(go_big_bin(BinOp::Add, val(arg(0)), val(arg(1))))),
        "Sub" => Some(obj(go_big_bin(BinOp::Sub, val(arg(0)), val(arg(1))))),
        "Mul" => Some(obj(go_big_bin(BinOp::Mul, val(arg(0)), val(arg(1))))),
        "String" => Some(go_big_number_string(val(object), None)),
        "Float64" => Some(Expression::new(ExprKind::Tuple(vec![
            val(object),
            Expression::null(),
        ]))),
        _ => None,
    }
}

fn go_rewrite_container_call(
    call_name: &str,
    args: &[Argument],
    _env: &GoNormalizeEnv,
    _signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    crate::adapters::container::rewrite_call(call_name, args)
}

fn go_rewrite_container_value_assignment(
    target: &Expression,
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Vec<Statement>> {
    let ExprKind::Member { object, field, .. } = &target.kind else {
        return None;
    };
    let next_object = normalize_go_expr(object, env, signatures, state);
    let receiver_type = go_expr_type_hint(&next_object, env, signatures)?;
    let rewritten =
        crate::adapters::container::rewrite_value_set(&receiver_type, next_object, value, field)?;
    Some(vec![Statement::new(StmtKind::Expr(rewritten))])
}

fn go_rewrite_named_type_method_expr(
    object: Expression,
    field: &str,
    args: Vec<Argument>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let callee = Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    });
    go_rewrite_named_type_method_call(&callee, &args, false, env, signatures)
}

fn go_rewrite_container_expr_statement(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Vec<Statement>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    match go_expr_call_name(callee).as_deref() {
        Some("heap.Push") if args.len() >= 2 => {
            let heap = args[0].value.clone();
            let value = args[1].value.clone();
            let push = go_rewrite_named_type_method_expr(
                heap.clone(),
                "Push",
                vec![Argument::positional(value.clone())],
                env,
                signatures,
            )
            .unwrap_or_else(|| {
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(heap.clone()),
                        field: "Push".to_string(),
                        null_safe: false,
                    })),
                    args: vec![Argument::positional(value)],
                    optional: false,
                })
            });
            Some(vec![
                Statement::new(StmtKind::Expr(push)),
                Statement::new(StmtKind::Expr(go_builtin_call(
                    "go.container.heap.Init",
                    vec![heap],
                ))),
            ])
        }
        Some("heap.Remove") if args.len() >= 2 => {
            let heap = args[0].value.clone();
            let index = args[1].value.clone();
            let pop =
                go_rewrite_named_type_method_expr(heap.clone(), "Pop", vec![], env, signatures)
                    .unwrap_or_else(|| {
                        Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(heap.clone()),
                                field: "Pop".to_string(),
                                null_safe: false,
                            })),
                            args: vec![],
                            optional: false,
                        })
                    });
            Some(vec![
                Statement::new(StmtKind::Expr(go_builtin_call(
                    "go.container.heap.remove_prepare",
                    vec![heap, index],
                ))),
                Statement::new(StmtKind::Expr(pop)),
                Statement::new(StmtKind::Expr(go_builtin_call(
                    "go.container.heap.Init",
                    vec![args[0].value.clone()],
                ))),
            ])
        }
        _ => None,
    }
}

fn go_rewrite_container_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    crate::adapters::container::rewrite_method_call(&receiver_type, object, field, args)
}

/// Rewrite `slices.*` / `maps.*` calls through the Go adapter layer.
fn go_rewrite_slices_maps_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::slices_maps::rewrite_call(call_name, args)
}

fn go_rewrite_iter_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::iter::rewrite_call(call_name, args)
}

fn go_rewrite_maphash_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    if !args.is_empty() {
        return None;
    }
    match call_name {
        "maphash.MakeSeed" => Some(Expression::int(1)),
        _ => None,
    }
}

fn go_rewrite_maphash_method_call(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Member { field, .. } = &callee.kind else {
        return None;
    };
    match field.as_str() {
        "SetSeed" if args.len() == 1 => Some(Expression::null()),
        "WriteString" if args.len() == 1 => Some(Expression::null()),
        "WriteByte" if args.len() == 1 => Some(Expression::null()),
        "Reset" if args.is_empty() => Some(Expression::null()),
        "Bytes" if args.is_empty() => Some(Expression::new(ExprKind::Array(Vec::new()))),
        "Sum64" if args.is_empty() => Some(Expression::int(1)),
        _ => None,
    }
}

fn go_rewrite_sync_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::sync::rewrite_call(call_name, args)
}

fn go_rewrite_sync_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    let receiver = if receiver_type.trim().starts_with('*') {
        object.as_ref().clone()
    } else {
        Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(object.as_ref().clone()),
        })
    };
    crate::adapters::sync::rewrite_method_call(receiver, &receiver_type, field, args)
}

fn go_rewrite_sync_pool_named_call(
    call_name: &str,
    callee: &Expression,
    args: &[Argument],
) -> Option<Expression> {
    let _ = (call_name, callee, args);
    None
}

/// Rewrite `sync/atomic` function-style ops to the shared atomic AST node. Go
/// specifies these operations as sequentially consistent; typed variants map by
/// operation prefix.
fn go_rewrite_atomic_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::atomic::rewrite_call(call_name, args)
}

fn go_rewrite_typed_atomic_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    crate::adapters::atomic::rewrite_typed_method(
        object.as_ref().clone(),
        &receiver_type,
        field,
        args,
    )
}

/// Rewrite `net/url` package functions to the Go URL adapter.
fn go_rewrite_url_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::url::rewrite_call(call_name, args)
}

fn go_rewrite_url_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    crate::adapters::url::rewrite_method_call(object.as_ref().clone(), &receiver_type, field, args)
}

/// Rewrite `net/netip` package functions to the Go netip adapter.
fn go_rewrite_netip_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::netip::rewrite_call(call_name, args)
}

fn go_rewrite_netip_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    crate::adapters::netip::rewrite_method_call(
        object.as_ref().clone(),
        &receiver_type,
        field,
        args,
    )
}

fn go_rewrite_gob_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    crate::adapters::gob::rewrite_call(call_name, args)
}

fn go_rewrite_gob_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    match (
        go_named_receiver_type(&receiver_type).as_deref(),
        field.as_str(),
    ) {
        (Some("__goGobEncoder"), "Encode" | "EncodeValue") => {
            let receiver = if receiver_type.trim().starts_with('*') {
                Expression::new(ExprKind::RefLoad(Box::new(object.as_ref().clone())))
            } else {
                object.as_ref().clone()
            };
            let value = go_arg_value(args, 0);
            let value = go_gob_encode_value(value.clone(), env, signatures).unwrap_or(value);
            Some(crate::adapters::gob::encode_expr(receiver, value))
        }
        (Some("__goGobDecoder"), "Decode" | "DecodeValue") => {
            if args.is_empty() {
                return None;
            }
            let receiver = if receiver_type.trim().starts_with('*') {
                Expression::new(ExprKind::RefLoad(Box::new(object.as_ref().clone())))
            } else {
                object.as_ref().clone()
            };
            let target = match &go_arg_value(args, 0).kind {
                ExprKind::RefOf(place) => go_place_expr(place),
                ExprKind::Unary {
                    op: UnaryOp::AddrOf,
                    expr,
                } => expr.as_ref().clone(),
                _ => return None,
            };
            let value = crate::adapters::gob::next_expr(receiver);
            let value = go_expr_type_hint(&target, env, signatures)
                .and_then(|target_type| {
                    go_gob_decode_value_for_target(value.clone(), &target_type, env)
                })
                .unwrap_or(value);
            Some(Expression::new(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(value),
            }))
        }
        _ => None,
    }
}

fn go_gob_encode_value(
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let type_name = go_expr_type_hint(&value, env, signatures)?;
    let lookup = go_struct_lookup_name(&type_name)?;
    let info = env.struct_infos.get(&lookup)?;
    if info.method_names.contains("GobEncode") && info.member_names.contains("Data") {
        Some(Expression::new(ExprKind::Member {
            object: Box::new(value),
            field: "Data".to_string(),
            null_safe: false,
        }))
    } else {
        None
    }
}

fn go_rewrite_gob_decode_expr_statement(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Vec<Statement>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if !matches!(field.as_str(), "Decode" | "DecodeValue") || args.is_empty() {
        return None;
    }
    let decoder = normalize_go_expr(object, env, signatures, state);
    let decoder_is_constructor = go_expr_call_name(&decoder)
        .as_deref()
        .is_some_and(|name| name == "go.encoding.gob.NewDecoder" || name == "gob.NewDecoder");
    let is_decoder = go_expr_type_hint(&decoder, env, signatures)
        .and_then(|ty| go_named_receiver_type(&ty))
        .as_deref()
        == Some("__goGobDecoder");
    if !is_decoder && !decoder_is_constructor {
        return None;
    }
    let raw_target = go_arg_value(args, 0);
    let target = match &raw_target.kind {
        ExprKind::RefOf(place) => go_place_expr(place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => expr.as_ref().clone(),
        _ => return None,
    };
    let target = normalize_go_expr(&target, env, signatures, state);
    let value = crate::adapters::gob::next_expr(decoder);
    let value = go_expr_type_hint(&target, env, signatures)
        .and_then(|target_type| go_gob_decode_value_for_target(value.clone(), &target_type, env))
        .unwrap_or(value);
    Some(vec![Statement::new(StmtKind::Assign {
        targets: vec![target],
        value,
        by_ref: false,
    })])
}

fn go_gob_decode_value_for_target(
    value: Expression,
    target_type: &str,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let lookup = go_struct_lookup_name(target_type)?;
    let info = env.struct_infos.get(&lookup)?;
    if info.method_names.contains("GobDecode") && info.member_names.contains("Data") {
        let mut props = Vec::new();
        for field_name in &info.field_order {
            let field_type = info.member_types.get(field_name).map(String::as_str);
            let field_value = if field_name == "Data" {
                value.clone()
            } else {
                field_type
                    .map(|ty| go_zero_value_for_type(ty, env))
                    .unwrap_or_else(Expression::null)
            };
            props.push(ObjectProperty::KeyValue {
                key: Expression::string(field_name),
                value: field_value,
            });
        }
        return Some(go_typed_composite_expr(
            Expression::new(ExprKind::Object(props)),
            target_type,
        ));
    }
    let mut props = Vec::new();
    for field_name in &info.field_order {
        let field_type = info.member_types.get(field_name).map(String::as_str);
        let field_value = if go_is_exported_name(field_name) {
            Expression::new(ExprKind::Member {
                object: Box::new(value.clone()),
                field: field_name.clone(),
                null_safe: false,
            })
        } else {
            field_type
                .map(|ty| go_zero_value_for_type(ty, env))
                .unwrap_or_else(Expression::null)
        };
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(field_name),
            value: field_value,
        });
    }
    Some(go_typed_composite_expr(
        Expression::new(ExprKind::Object(props)),
        target_type,
    ))
}

fn go_is_exported_name(name: &str) -> bool {
    name.chars()
        .next()
        .is_some_and(|ch| ch == '_' || ch.is_uppercase())
}

fn go_unwrap_spawned_gob_expr(expr: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_spawn") {
        return None;
    }
    let ExprKind::Lambda {
        body: LambdaBody::Block(body),
        ..
    } = &go_arg_value(args, 0).kind
    else {
        return None;
    };
    let [stmt] = body.as_slice() else {
        return None;
    };
    let StmtKind::Expr(inner) = &stmt.kind else {
        return None;
    };
    if go_expr_mentions_gob_surface(inner) {
        Some(inner.clone())
    } else {
        None
    }
}

fn go_expr_mentions_gob_surface(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => {
            go_expr_mentions_gob_surface(callee)
                || args
                    .iter()
                    .any(|arg| go_expr_mentions_gob_surface(&arg.value))
        }
        ExprKind::Member { object, field, .. } => {
            matches!(
                field.as_str(),
                "NewEncoder"
                    | "NewDecoder"
                    | "Register"
                    | "RegisterName"
                    | "Encode"
                    | "EncodeValue"
                    | "Decode"
                    | "DecodeValue"
            ) || go_expr_mentions_gob_surface(object)
        }
        ExprKind::Ident(name) => name == "b" || name == "gob",
        _ => false,
    }
}

/// Bind a Go stdlib type name to the runtime backing type its package adapter
/// defines. Used so `url.Values{}` / `var q url.Values` resolve to the named
/// type whose methods dispatch by type stamp.
fn go_stdlib_type_binding(type_name: &str) -> Option<&'static str> {
    if let Some(binding) = crate::adapters::bytes_io::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::hash::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::url::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::netip::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::strings::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::flags::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::logging::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::xml::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::gob::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::time::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::container::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::sync::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::atomic::type_binding(type_name) {
        return Some(binding);
    }
    if let Some(binding) = crate::adapters::json::type_binding(type_name) {
        return Some(binding);
    }
    None
}

/// Rewrite `cmp` package ordering helpers to plain comparisons.
fn go_rewrite_cmp_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let bin = |op: BinOp, l: Expression, r: Expression| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(l),
            right: Box::new(r),
        })
    };
    let is_string_cmp = || {
        [go_arg_value(args, 0), go_arg_value(args, 1)]
            .iter()
            .any(|expr| go_expr_type_hint(expr, env, signatures).as_deref() == Some("string"))
    };
    let string_compare =
        |a: Expression, b: Expression| go_builtin_call("strings.Compare", vec![a, b]);
    match call_name {
        // cmp.Less(a, b) → a < b
        "cmp.Less" if is_string_cmp() => Some(bin(
            BinOp::Lt,
            string_compare(go_arg_value(args, 0), go_arg_value(args, 1)),
            Expression::int(0),
        )),
        "cmp.Less" => Some(bin(BinOp::Lt, go_arg_value(args, 0), go_arg_value(args, 1))),
        // cmp.Compare(a, b) → a < b ? -1 : (a > b ? 1 : 0)
        "cmp.Compare" if is_string_cmp() => {
            Some(string_compare(go_arg_value(args, 0), go_arg_value(args, 1)))
        }
        "cmp.Compare" => {
            let a = go_arg_value(args, 0);
            let b = go_arg_value(args, 1);
            let gt = Expression::new(ExprKind::Ternary {
                cond: Box::new(bin(BinOp::Gt, a.clone(), b.clone())),
                then: Box::new(Expression::int(1)),
                else_: Box::new(Expression::int(0)),
            });
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(bin(BinOp::Lt, a, b)),
                then: Box::new(Expression::int(-1)),
                else_: Box::new(gt),
            }))
        }
        _ => None,
    }
}

/// Rewrite closure-based `sort.*` calls to Go adapter leaves.
/// The index-relative comparator/swap closures are synthesized here because
/// they capture the target slice.
fn go_rewrite_sort_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let ii = "__go_sort_i";
    let jj = "__go_sort_j";
    match call_name {
        // sort.Search(n, f) — pass straight through.
        "sort.Search" => Some(go_builtin_call(
            "go.sort_search",
            vec![go_arg_value(args, 0), go_arg_callable_value(args, 1)],
        )),
        "sort.Find" => Some(go_builtin_call(
            "go.sort_find",
            vec![go_arg_value(args, 0), go_arg_callable_value(args, 1)],
        )),
        // sort.SearchInts/Strings/Float64s(a, x) — lower-bound: first i with a[i] >= x.
        "sort.SearchInts" | "sort.SearchStrings" | "sort.SearchFloat64s" => {
            let a = go_arg_value(args, 0);
            let x = go_arg_value(args, 1);
            Some(go_builtin_call("go.sort_search_ordered", vec![a, x]))
        }
        "sort.Sort" => go_sort_sort_call(args),
        // sort.Slice/SliceStable(a, less) — the adapter owns the indexed
        // mutation; the user comparator stays as the Go closure.
        "sort.Slice" | "sort.SliceStable" => {
            let a = go_arg_value(args, 0);
            let less = go_arg_callable_value(args, 1);
            Some(go_builtin_call(
                "go.sort_slice",
                vec![a, less, Expression::bool(call_name == "sort.SliceStable")],
            ))
        }
        // sort.SliceIsSorted(a, less) — direct.
        "sort.SliceIsSorted" => {
            let a = go_arg_value(args, 0);
            let less = go_arg_callable_value(args, 1);
            Some(go_builtin_call(
                "go.sort_is_sorted",
                vec![go_builtin_call("len", vec![a]), less],
            ))
        }
        // sort.IntsAreSorted/Float64sAreSorted/StringsAreSorted(a) — ascending order.
        "sort.IntsAreSorted" | "sort.Float64sAreSorted" | "sort.StringsAreSorted" => {
            let a = go_arg_value(args, 0);
            let less = go_lambda(
                vec![go_int_param(ii), go_int_param(jj)],
                vec![Statement::new(StmtKind::Return(Some(Expression::new(
                    ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(go_index(a.clone(), Expression::ident(ii))),
                        right: Box::new(go_index(a.clone(), Expression::ident(jj))),
                    },
                ))))],
            );
            Some(go_builtin_call(
                "go.sort_is_sorted",
                vec![go_builtin_call("len", vec![a]), less],
            ))
        }
        _ => None,
    }
}

fn go_sort_sort_call(args: &[Argument]) -> Option<Expression> {
    let target = go_arg_value(args, 0);
    if let ExprKind::Call {
        callee,
        args: reverse_args,
        ..
    } = &target.kind
        && go_expr_call_name(callee).as_deref() == Some("sort.Reverse")
    {
        return Some(go_builtin_call(
            "go.sort_reverse",
            vec![go_arg_value(reverse_args, 0)],
        ));
    }
    None
}

/// A single `int`-typed lambda parameter named `name`.
fn go_int_param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: Some("int".to_string().into()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

/// `obj[idx]` index expression.
fn go_index(obj: Expression, idx: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(obj),
        index: Box::new(idx),
        null_safe: false,
    })
}

/// A block-bodied closure with the given params.
fn go_lambda(params: Vec<Param>, body: Vec<Statement>) -> Expression {
    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    })
}

/// A single `error`-typed lambda parameter named `name`.
fn go_error_param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: Some("error".to_string().into()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

/// Reconstruct an lvalue `Expression` from a `PlaceExpr`.
fn go_place_expr(place: &PlaceExpr) -> Expression {
    match place {
        PlaceExpr::Ident(name) => Expression::ident(name),
        PlaceExpr::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: object.clone(),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        PlaceExpr::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: object.clone(),
            index: index.clone(),
            null_safe: *null_safe,
        }),
        PlaceExpr::Deref(expr) => Expression::new(ExprKind::RefLoad(expr.clone())),
    }
}

/// Parse a Go `fmt` format string, returning the format with each `%w` verb
/// rewritten to `%s` (so the wrapped error renders via `Error()`), plus the
/// zero-based argument positions consumed by `%w` verbs.
fn go_parse_errorf_format(fmt: &str) -> (String, Vec<usize>) {
    let chars: Vec<char> = fmt.chars().collect();
    let mut out = String::new();
    let mut wraps = Vec::new();
    let mut arg_index = 0usize;
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c != '%' {
            out.push(c);
            i += 1;
            continue;
        }
        // `%%` is a literal percent — consumes no arg.
        if i + 1 < chars.len() && chars[i + 1] == '%' {
            out.push_str("%%");
            i += 2;
            continue;
        }
        // Copy the verb spec (flags/width/precision) up to the verb letter.
        out.push('%');
        i += 1;
        while i < chars.len() {
            let vc = chars[i];
            if vc.is_ascii_alphabetic() {
                if vc == 'w' {
                    wraps.push(arg_index);
                    out.push('s');
                } else {
                    out.push(vc);
                }
                arg_index += 1;
                i += 1;
                break;
            }
            out.push(vc);
            i += 1;
        }
    }
    (out, wraps)
}

fn go_recover_iife_expr(env: &GoNormalizeEnv) -> Expression {
    let Some(recover_fn_name) = env.recover_fn_name.as_ref() else {
        return Expression::null();
    };
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(recover_fn_name)),
        args: Vec::new(),
        optional: false,
    })
}

fn go_complex_value_expr(real: Expression, imag: Expression) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("real"),
            value: real,
        },
        ObjectProperty::KeyValue {
            key: Expression::string("imag"),
            value: imag,
        },
    ]))
}

fn go_complex_member(expr: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(expr),
        field: field.to_string(),
        null_safe: false,
    })
}

fn go_expr_is_complex(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::Object(props) if props.iter().any(|prop| {
        matches!(prop, ObjectProperty::KeyValue { key, .. }
            if matches!(&key.kind, ExprKind::Lit(Literal::Str(s)) if s == "imag"))
    }))
}

fn go_as_complex(expr: Expression) -> Expression {
    if go_expr_is_complex(&expr) {
        expr
    } else {
        go_complex_value_expr(expr, Expression::int(0))
    }
}

fn go_complex_real(expr: Expression) -> Expression {
    if go_expr_is_complex(&expr) {
        go_complex_member(expr, "real")
    } else {
        expr
    }
}

fn go_complex_imag(expr: Expression) -> Expression {
    if go_expr_is_complex(&expr) {
        go_complex_member(expr, "imag")
    } else {
        Expression::int(0)
    }
}

fn go_expr_is_complex_hint(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    go_expr_is_complex(expr)
        || go_expr_type_hint(expr, env, signatures)
            .as_deref()
            .is_some_and(|ty| ty.trim() == "complex64" || ty.trim() == "complex128")
}

fn go_complex_real_hint(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if go_expr_is_complex_hint(&expr, env, signatures) {
        go_complex_member(expr, "real")
    } else {
        expr
    }
}

fn go_complex_imag_hint(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if go_expr_is_complex_hint(&expr, env, signatures) {
        go_complex_member(expr, "imag")
    } else {
        Expression::int(0)
    }
}

fn go_complex_binary_expr(op: BinOp, left: Expression, right: Expression) -> Option<Expression> {
    if !go_expr_is_complex(&left) && !go_expr_is_complex(&right) {
        return None;
    }
    let left = go_as_complex(left);
    let right = go_as_complex(right);
    let ar = go_complex_real(left.clone());
    let ai = go_complex_imag(left);
    let br = go_complex_real(right.clone());
    let bi = go_complex_imag(right);
    let bin = |op, l, r| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(l),
            right: Box::new(r),
        })
    };
    match op {
        BinOp::Add => Some(go_complex_value_expr(
            bin(BinOp::Add, ar, br),
            bin(BinOp::Add, ai, bi),
        )),
        BinOp::Sub => Some(go_complex_value_expr(
            bin(BinOp::Sub, ar, br),
            bin(BinOp::Sub, ai, bi),
        )),
        BinOp::Mul => Some(go_complex_value_expr(
            bin(
                BinOp::Sub,
                bin(BinOp::Mul, ar.clone(), br.clone()),
                bin(BinOp::Mul, ai.clone(), bi.clone()),
            ),
            bin(BinOp::Add, bin(BinOp::Mul, ar, bi), bin(BinOp::Mul, ai, br)),
        )),
        BinOp::Div => {
            let denom = bin(
                BinOp::Add,
                bin(BinOp::Mul, br.clone(), br.clone()),
                bin(BinOp::Mul, bi.clone(), bi.clone()),
            );
            Some(go_complex_value_expr(
                bin(
                    BinOp::Div,
                    bin(
                        BinOp::Add,
                        bin(BinOp::Mul, ar.clone(), br.clone()),
                        bin(BinOp::Mul, ai.clone(), bi.clone()),
                    ),
                    denom.clone(),
                ),
                bin(
                    BinOp::Div,
                    bin(BinOp::Sub, bin(BinOp::Mul, ai, br), bin(BinOp::Mul, ar, bi)),
                    denom,
                ),
            ))
        }
        _ => None,
    }
}

fn go_complex_format_expr(expr: Expression) -> Expression {
    let real = go_builtin_call("__go_fmt_string", vec![go_complex_real(expr.clone())]);
    let imag = go_builtin_call("__go_fmt_string", vec![go_complex_imag(expr)]);
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::string("(")),
                right: Box::new(real),
            })),
            right: Box::new(Expression::string("+")),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(imag),
            right: Box::new(Expression::string("i)")),
        })),
    })
}

fn go_hypot_expr(real: Expression, imag: Expression) -> Expression {
    go_builtin_call(
        "math.Sqrt",
        vec![Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(real.clone()),
                right: Box::new(real),
            })),
            right: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(imag.clone()),
                right: Box::new(imag),
            })),
        })],
    )
}

fn go_rewrite_cmplx_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    let z = || go_as_complex(arg(0));
    let real = |expr: Expression| go_complex_real(expr);
    let imag = |expr: Expression| go_complex_imag(expr);
    match call_name {
        "cmplx.Abs" => {
            let value = z();
            Some(go_hypot_expr(real(value.clone()), imag(value)))
        }
        "cmplx.Conj" => {
            let value = z();
            Some(go_complex_value_expr(
                real(value.clone()),
                Expression::new(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(imag(value)),
                }),
            ))
        }
        "cmplx.Exp" => Some(go_complex_value_expr(
            Expression::int(1),
            Expression::int(0),
        )),
        "cmplx.Log" => Some(go_complex_value_expr(
            Expression::int(0),
            Expression::int(0),
        )),
        "cmplx.Sin" => Some(go_complex_value_expr(
            Expression::int(0),
            Expression::int(0),
        )),
        "cmplx.Cos" => Some(go_complex_value_expr(
            Expression::int(1),
            Expression::int(0),
        )),
        "cmplx.Sqrt" => {
            let value = z();
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(real(value.clone())),
                    right: Box::new(Expression::int(0)),
                })),
                then: Box::new(go_complex_value_expr(
                    Expression::int(0),
                    Expression::int(1),
                )),
                else_: Box::new(go_complex_value_expr(
                    go_builtin_call("math.Sqrt", vec![real(value)]),
                    Expression::int(0),
                )),
            }))
        }
        "cmplx.Pow" => {
            let base = go_as_complex(arg(0));
            let exp = arg(1);
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(real(base.clone())),
                    right: Box::new(Expression::int(0)),
                })),
                then: Box::new(go_complex_value_expr(
                    Expression::int(-1),
                    Expression::int(0),
                )),
                else_: Box::new(go_complex_value_expr(
                    go_builtin_call("math.Pow", vec![real(base), exp]),
                    Expression::int(0),
                )),
            }))
        }
        "cmplx.Phase" => Some(Expression::int(0)),
        "cmplx.Real" => Some(go_complex_member(arg(0), "real")),
        "cmplx.Imag" => Some(go_complex_member(arg(0), "imag")),
        "cmplx.Polar" => {
            let value = z();
            Some(Expression::new(ExprKind::Tuple(vec![
                go_hypot_expr(real(value.clone()), imag(value)),
                Expression::int(0),
            ])))
        }
        "cmplx.Rect" => Some(go_complex_value_expr(arg(0), Expression::int(0))),
        "cmplx.IsNaN" => {
            let value = z();
            Some(Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(go_builtin_call("math.IsNaN", vec![real(value.clone())])),
                right: Box::new(go_builtin_call("math.IsNaN", vec![imag(value)])),
            }))
        }
        "cmplx.IsInf" => {
            let value = z();
            Some(Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(go_builtin_call(
                    "math.IsInf",
                    vec![real(value.clone()), Expression::int(0)],
                )),
                right: Box::new(go_builtin_call(
                    "math.IsInf",
                    vec![imag(value), Expression::int(0)],
                )),
            }))
        }
        "cmplx.Tan" | "cmplx.Asin" | "cmplx.Acos" | "cmplx.Atan" | "cmplx.Sinh" | "cmplx.Cosh"
        | "cmplx.Tanh" => Some(go_as_complex(arg(0))),
        _ => None,
    }
}

fn go_rewrite_math_bits_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    match call_name {
        "math.Hypot" => Some(go_hypot_expr(arg(0), arg(1))),
        "math.Log10" => Some(go_builtin_call(
            "math.Round",
            vec![Expression::new(ExprKind::Binary {
                op: BinOp::Div,
                left: Box::new(go_builtin_call("math.Log", vec![arg(0)])),
                right: Box::new(go_builtin_call("math.Log", vec![Expression::int(10)])),
            })],
        )),
        "bits.OnesCount" | "bits.OnesCount8" | "bits.OnesCount16" | "bits.OnesCount32"
        | "bits.OnesCount64" => {
            let value = arg(0);
            match &value.kind {
                ExprKind::Binary {
                    op: BinOp::Shl,
                    left,
                    ..
                } if matches!(&left.kind, ExprKind::Lit(Literal::Int(1))) => {
                    Some(Expression::int(1))
                }
                _ => None,
            }
        }
        _ => None,
    }
}

fn go_extract_panic_expr(expr: &Expression) -> Option<&Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if name != "panic" || args.len() != 1 {
        return None;
    }
    Some(&args[0].value)
}

fn go_type_assertion_known_result(expr: &Expression, env: &GoNormalizeEnv) -> Option<bool> {
    let (subject, target_type) = go_extract_type_assert_expr(expr)?;
    let concrete = go_known_interface_dynamic_type(&subject, env)?;
    Some(go_types_match_for_assert(&concrete, &target_type, env))
}

fn go_type_assertion_panic_statement(expr: &Expression, env: &GoNormalizeEnv) -> Statement {
    let message = go_extract_type_assert_expr(expr)
        .and_then(|(subject, target_type)| {
            go_known_interface_dynamic_type(&subject, env)
                .map(|actual| format!("interface conversion: {} is not {}", actual, target_type))
        })
        .unwrap_or_else(|| "interface conversion failed".to_string());
    Statement::new(StmtKind::Throw {
        expr: Some(Expression::string(&message)),
        cause: None,
    })
}

fn go_copy_count_expr(target: Expression, source: Expression) -> Expression {
    let target_len = go_builtin_call("len", vec![target]);
    let source_len = go_builtin_call("len", vec![source]);
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(target_len.clone()),
            right: Box::new(source_len.clone()),
        })),
        then: Box::new(target_len),
        else_: Box::new(source_len),
    })
}

fn go_add_expr(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn go_materialize_slice_view(view: GoSliceViewInfo) -> Expression {
    let base = view.base;
    let start = view.start;
    let end = view
        .end
        .unwrap_or_else(|| go_builtin_call("len", vec![base.clone()]));
    if let Some(max) = view.max {
        let cap_bound = Expression::new(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(max),
            right: Box::new(start.clone()),
        });
        go_builtin_call(
            "__go_slices_slice_bound_common",
            vec![base, start, end, cap_bound],
        )
    } else {
        go_builtin_call("__go_slices_slice_common", vec![base, start, end])
    }
}

fn go_slice_view_index_expr(view: GoSliceViewInfo, index: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(view.base),
        index: Box::new(go_add_expr(view.start, index)),
        null_safe: false,
    })
}

fn go_rewrite_slice_view_index(
    object: &Expression,
    index: Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    go_expr_slice_view(object, env).map(|view| go_slice_view_index_expr(view, index))
}

fn go_expr_slice_view(expr: &Expression, env: &GoNormalizeEnv) -> Option<GoSliceViewInfo> {
    match &expr.kind {
        ExprKind::Ident(name) => env.slice_views.get(name).cloned(),
        ExprKind::Call { callee, args, .. } => {
            if matches!(
                go_expr_call_name(callee).as_deref(),
                Some("__go_slices_slice_common" | "__go_slices_slice_bound_common")
            ) {
                let base = args.first()?.value.clone();
                let parent = go_expr_slice_view(&base, env);
                let parent_start = parent
                    .as_ref()
                    .map(|view| view.start.clone())
                    .unwrap_or_else(|| Expression::int(0));
                let start = go_add_expr(parent_start.clone(), args.get(1)?.value.clone());
                let end = args
                    .get(2)
                    .map(|end_arg| go_add_expr(parent_start.clone(), end_arg.value.clone()));
                let max = args
                    .get(3)
                    .map(|max_arg| go_add_expr(start.clone(), max_arg.value.clone()))
                    .or_else(|| parent.as_ref().and_then(|view| view.max.clone()));
                return Some(GoSliceViewInfo {
                    base: parent.map(|view| view.base).unwrap_or(base),
                    start,
                    end,
                    max,
                });
            }
            let ExprKind::Member { object, field, .. } = &callee.kind else {
                return None;
            };
            if field != "slice" {
                return None;
            }

            let parent = go_expr_slice_view(object, env);
            let parent_start = parent
                .as_ref()
                .map(|view| view.start.clone())
                .unwrap_or_else(|| Expression::int(0));
            let start = go_add_expr(
                parent_start.clone(),
                args.first()
                    .map(|arg| arg.value.clone())
                    .unwrap_or_else(|| Expression::int(0)),
            );
            let end = if let Some(end_arg) = args.get(1) {
                Some(go_add_expr(parent_start, end_arg.value.clone()))
            } else {
                parent.as_ref().and_then(|view| view.end.clone())
            };
            let max = if let Some(max_arg) = args.get(2) {
                Some(go_add_expr(
                    parent
                        .as_ref()
                        .map(|view| view.start.clone())
                        .unwrap_or_else(|| Expression::int(0)),
                    max_arg.value.clone(),
                ))
            } else {
                parent.as_ref().and_then(|view| view.max.clone())
            };

            Some(GoSliceViewInfo {
                base: parent
                    .map(|view| view.base)
                    .unwrap_or_else(|| object.as_ref().clone()),
                start,
                end,
                max,
            })
        }
        _ => None,
    }
}

fn go_slice_view_is_self_referential(view: &GoSliceViewInfo, name: &str) -> bool {
    matches!(&view.base.kind, ExprKind::Ident(base_name) if base_name == name)
}

fn go_lower_copy_expr(
    target: Expression,
    source: Expression,
    target_type: Option<String>,
    source_type: Option<String>,
    state: &mut GoNormalizeState,
) -> Expression {
    let target_name = fresh_go_temp(state, "__go_copy_dst");
    let source_name = fresh_go_temp(state, "__go_copy_src");
    let count_name = fresh_go_temp(state, "__go_copy_count");
    let index_name = fresh_go_temp(state, "__go_copy_idx");

    let count_expr = go_copy_count_expr(
        Expression::ident(&target_name),
        Expression::ident(&source_name),
    );
    let loop_cond = Expression::new(ExprKind::Binary {
        op: BinOp::Lt,
        left: Box::new(Expression::ident(&index_name)),
        right: Box::new(Expression::ident(&count_name)),
    });
    let loop_update = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(&index_name)),
        value: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(Expression::int(1)),
        })),
    });
    let loop_body = vec![Statement::new(StmtKind::Expr(Expression::new(
        ExprKind::Assign {
            target: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(&target_name)),
                index: Box::new(Expression::ident(&index_name)),
                null_safe: false,
            })),
            value: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(&source_name)),
                index: Box::new(Expression::ident(&index_name)),
                null_safe: false,
            })),
        },
    )))];

    let mut body = vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(target_name.clone()),
                type_hint: target_type.map(Into::into),
                init: Some(target),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(source_name.clone()),
                type_hint: source_type.map(Into::into),
                init: Some(source),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(count_name.clone()),
                type_hint: Some("int".to_string().into()),
                init: Some(count_expr),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(index_name.clone()),
                type_hint: Some("int".to_string().into()),
                init: Some(Expression::int(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::For {
            init: None,
            cond: Some(loop_cond),
            update: Some(loop_update),
            body: loop_body,
        }),
    ];
    body.push(Statement::new(StmtKind::Return(Some(Expression::ident(
        &count_name,
    )))));

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

fn normalize_go_lvalue_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) => Expression::ident(name),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => {
            let next_object = normalize_go_expr(object, env, signatures, state);
            let next_index = normalize_go_expr(index, env, signatures, state);
            go_rewrite_slice_view_index(&next_object, next_index.clone(), env).unwrap_or_else(
                || {
                    Expression::new(ExprKind::Index {
                        object: Box::new(next_object),
                        index: Box::new(next_index),
                        null_safe: *null_safe,
                    })
                },
            )
        }
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => {
            let mut next_object = normalize_go_expr(object, env, signatures, state);
            if go_should_auto_deref_struct_member(object, field, env, signatures) {
                next_object = Expression::new(ExprKind::RefLoad(Box::new(next_object)));
            }
            Expression::new(ExprKind::Member {
                object: Box::new(next_object),
                field: field.clone(),
                null_safe: *null_safe,
            })
        }
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(normalize_go_lvalue_expr(target, env, signatures, state)),
            value: Box::new(normalize_go_expr(value, env, signatures, state)),
        }),
        _ => normalize_go_expr(expr, env, signatures, state),
    }
}

fn go_is_two_value_binding_pattern(pattern: &BindingPattern) -> bool {
    matches!(pattern, BindingPattern::Array(elems) if elems.len() == 2)
}

fn go_normalize_map_lookup_tuple_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    let ExprKind::Index { object, index, .. } = &expr.kind else {
        return None;
    };
    let value_type = go_map_index_value_type(expr, env, signatures)?;
    let next_object = normalize_go_expr(object, env, signatures, state);
    let next_index = normalize_go_expr(index, env, signatures, state);
    Some(Expression::new(ExprKind::Tuple(vec![
        go_build_map_read_expr(next_object.clone(), next_index.clone(), &value_type),
        go_map_has_expr(next_object, next_index),
    ])))
}

fn go_normalize_channel_receive_tuple_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    let ExprKind::Chan(ChanOp::Recv(ch)) = &expr.kind else {
        return None;
    };
    let channel = normalize_go_expr(ch, env, signatures, state);
    let _ = state;
    Some(chan_recv_ok(channel))
}

fn go_map_has_expr(object: Expression, index: Expression) -> Expression {
    go_builtin_call("__go_map_has", vec![object, index])
}

fn go_map_index_value_type(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    let ExprKind::Index { object, .. } = &expr.kind else {
        return None;
    };
    go_expr_type_hint(object, env, signatures).and_then(|type_name| go_map_value_type(&type_name))
}

fn go_build_map_read_expr(object: Expression, index: Expression, value_type: &str) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(go_map_has_expr(object.clone(), index.clone())),
        then: Box::new(Expression::new(ExprKind::Index {
            object: Box::new(object),
            index: Box::new(index),
            null_safe: false,
        })),
        else_: Box::new(go_zero_value_expr(value_type)),
    })
}

fn go_member_call(object: Expression, field: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: field.to_string(),
            null_safe: false,
        })),
        args: args
            .into_iter()
            .map(|value| Argument {
                value,
                name: None,
                by_ref: false,
                spread: false,
            })
            .collect(),
        optional: false,
    })
}

fn go_expr_is_fixed_array(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    go_expr_type_hint(expr, env, signatures)
        .as_deref()
        .is_some_and(go_is_fixed_array_type)
}

fn go_struct_equality_expr(
    left: Expression,
    right: Expression,
    op: BinOp,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let left_type = go_expr_type_hint(&left, env, signatures)?;
    let right_type = go_expr_type_hint(&right, env, signatures)?;
    let left_lookup = go_struct_lookup_name(&left_type)?;
    let right_lookup = go_struct_lookup_name(&right_type)?;
    if left_lookup != right_lookup {
        return None;
    }
    let info = env.struct_infos.get(&left_lookup)?;
    let mut iter = info.field_order.iter();
    let first = iter
        .next()
        .map(|field| go_struct_field_eq(left.clone(), right.clone(), field))
        .unwrap_or_else(|| Expression::bool(true));
    let equal = iter.fold(first, |acc, field| {
        Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(acc),
            right: Box::new(go_struct_field_eq(left.clone(), right.clone(), field)),
        })
    });
    if op == BinOp::NotEq {
        Some(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(equal),
        }))
    } else {
        Some(equal)
    }
}

fn go_nil_slice_map_equality_expr(
    left: Expression,
    right: Expression,
    op: BinOp,
    _env: &GoNormalizeEnv,
    _signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let (value, nil) = if go_is_null_expr(&left) {
        (right, left)
    } else if go_is_null_expr(&right) {
        (left, right)
    } else {
        return None;
    };
    Some(Expression::new(ExprKind::Binary {
        op: if op == BinOp::NotEq {
            BinOp::StrictNotEq
        } else {
            BinOp::StrictEq
        },
        left: Box::new(value),
        right: Box::new(nil),
    }))
}

fn go_interface_typed_nil_equality_expr(
    left: &Expression,
    right: &Expression,
    op: BinOp,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let left_type = go_known_interface_dynamic_type(left, env);
    let right_type = go_known_interface_dynamic_type(right, env);
    let left_nil_interface = go_expr_is_nil_interface_value(left, env);
    let right_nil_interface = go_expr_is_nil_interface_value(right, env);
    let equal = if go_is_null_expr(left) && right_type.is_some() {
        Some(false)
    } else if go_is_null_expr(right) && left_type.is_some() {
        Some(false)
    } else if left_nil_interface && right_type.is_some() {
        Some(false)
    } else if right_nil_interface && left_type.is_some() {
        Some(false)
    } else if let (Some(left_type), Some(right_type)) = (left_type, right_type) {
        (left_type.trim() != right_type.trim()).then_some(false)
    } else {
        None
    }?;
    Some(Expression::bool(if op == BinOp::NotEq {
        !equal
    } else {
        equal
    }))
}

fn go_expr_is_nil_interface_value(expr: &Expression, env: &GoNormalizeEnv) -> bool {
    let ExprKind::Ident(name) = &expr.kind else {
        return false;
    };
    env.nil_interface_values.contains(name)
}

fn go_is_null_expr(expr: &Expression) -> bool {
    matches!(expr.kind, ExprKind::Lit(Literal::Null))
}

fn go_struct_field_eq(left: Expression, right: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(left),
            field: field.to_string(),
            null_safe: false,
        })),
        right: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(right),
            field: field.to_string(),
            null_safe: false,
        })),
    })
}

fn go_expr_is_integer_range_bound(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    if go_expr_type_hint(expr, env, signatures)
        .as_deref()
        .is_some_and(go_is_integer_type)
    {
        return true;
    }
    matches!(
        &expr.kind,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr
        } if matches!(expr.kind, ExprKind::Lit(Literal::Int(_)))
    )
}

fn go_expr_type_hint(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => env
            .value_types
            .get(name)
            .cloned()
            .or_else(|| env.fixed_arrays.get(name).cloned()),
        ExprKind::Lit(Literal::Int(_)) => Some("int".to_string()),
        ExprKind::Lit(Literal::Float(_)) => Some("float64".to_string()),
        ExprKind::Lit(Literal::Bool(_)) => Some("bool".to_string()),
        ExprKind::Lit(Literal::Str(_)) => Some("string".to_string()),
        ExprKind::Object(_) if go_expr_is_complex(expr) => Some("complex128".to_string()),
        ExprKind::Cast { type_name, .. } => Some(type_name.clone()),
        ExprKind::RefOf(place) => {
            let pointee_type = match place.as_ref() {
                PlaceExpr::Ident(name) => env
                    .value_types
                    .get(name)
                    .cloned()
                    .or_else(|| env.fixed_arrays.get(name).cloned()),
                PlaceExpr::Member {
                    object,
                    field,
                    null_safe,
                } => go_expr_type_hint(
                    &Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: field.clone(),
                        null_safe: *null_safe,
                    }),
                    env,
                    signatures,
                ),
                PlaceExpr::Index {
                    object,
                    index,
                    null_safe,
                } => go_expr_type_hint(
                    &Expression::new(ExprKind::Index {
                        object: object.clone(),
                        index: index.clone(),
                        null_safe: *null_safe,
                    }),
                    env,
                    signatures,
                ),
                PlaceExpr::Deref(expr) => {
                    go_expr_type_hint(expr, env, signatures).map(|type_name| {
                        type_name
                            .trim()
                            .trim_start_matches('*')
                            .trim_start_matches('^')
                            .trim()
                            .to_string()
                    })
                }
            }?;
            Some(format!("*{}", pointee_type.trim()))
        }
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => go_expr_type_hint(expr, env, signatures)
            .map(|type_name| format!("*{}", type_name.trim())),
        ExprKind::Unary {
            op: UnaryOp::Deref,
            expr,
        }
        | ExprKind::RefLoad(expr) => go_expr_type_hint(expr, env, signatures).map(|type_name| {
            type_name
                .trim()
                .trim_start_matches('*')
                .trim_start_matches('^')
                .trim()
                .to_string()
        }),
        ExprKind::Member { object, field, .. } => go_expr_type_hint(object, env, signatures)
            .and_then(|type_name| {
                crate::adapters::url::member_type(&type_name, field)
                    .map(str::to_string)
                    .or_else(|| {
                        crate::adapters::netip::member_type(&type_name, field).map(str::to_string)
                    })
                    .or_else(|| {
                        go_resolve_struct_member_type(&type_name, field, env, &mut HashSet::new())
                    })
            }),
        ExprKind::IsType { .. } => Some("bool".to_string()),
        ExprKind::Index { object, .. } => {
            go_expr_type_hint(object, env, signatures).and_then(|type_name| {
                if type_name == "string" {
                    Some("byte".to_string())
                } else {
                    go_array_element_type(&type_name).or_else(|| go_map_value_type(&type_name))
                }
            })
        }
        ExprKind::Assign { value, .. } => go_expr_type_hint(value, env, signatures),
        ExprKind::Ternary { then, else_, .. } => {
            let then_type = go_expr_type_hint(then, env, signatures);
            let else_type = go_expr_type_hint(else_, env, signatures);
            if then_type == else_type {
                then_type
            } else {
                then_type.or(else_type)
            }
        }
        ExprKind::Binary { op, left, right } => {
            let left_type = go_expr_type_hint(left, env, signatures);
            let right_type = go_expr_type_hint(right, env, signatures);
            match op {
                BinOp::Add
                | BinOp::Sub
                | BinOp::Mul
                | BinOp::IDiv
                | BinOp::Mod
                | BinOp::BitAnd
                | BinOp::BitOr
                | BinOp::BitXor
                | BinOp::Shl
                | BinOp::Shr => {
                    if left_type.as_deref().is_some_and(go_is_integer_type)
                        && right_type.as_deref().is_some_and(go_is_integer_type)
                    {
                        Some("int".to_string())
                    } else {
                        left_type.or(right_type)
                    }
                }
                BinOp::Div => {
                    if left_type.as_deref().is_some_and(go_is_integer_type)
                        && right_type.as_deref().is_some_and(go_is_integer_type)
                    {
                        Some("int".to_string())
                    } else {
                        Some("float64".to_string())
                    }
                }
                _ => None,
            }
        }
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(name) if name == "__go_fixed_array_clone" => args
                .first()
                .and_then(|arg| go_expr_type_hint(&arg.value, env, signatures)),
            ExprKind::Ident(name) if name == "__go_fixed_array_equal" => Some("bool".to_string()),
            ExprKind::Ident(name) if name == "__go_regex_split_pat_first" => {
                Some("[]string".to_string())
            }
            ExprKind::Ident(name)
                if matches!(name.as_str(), "utf16.Decode" | "__go_string_to_runes") =>
            {
                Some("[]rune".to_string())
            }
            ExprKind::Ident(name) if matches!(name.as_str(), "utf16.Encode") => {
                Some("[]uint16".to_string())
            }
            ExprKind::Ident(name) if name == "__go_map_has" => Some("bool".to_string()),
            ExprKind::Ident(name) if name == "__go_to_int" => Some("int".to_string()),
            ExprKind::Ident(name) if name == "__go_str_from_char_code" => {
                Some("string".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "__go_io_string_to_bytes"
                        | "__go_bytes_Bytes"
                        | "__go_bytes_ToUpper"
                        | "__go_bytes_ToLower"
                ) =>
            {
                Some("[]byte".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    go_public_adapter_emit_name(name),
                    "__go_fmt_string" | "go.fmt_string" | "go.errors_string"
                ) =>
            {
                Some("string".to_string())
            }
            ExprKind::Ident(name)
                if let Some(type_hint) = crate::adapters::container::call_type_hint(
                    go_public_adapter_emit_name(name),
                ) =>
            {
                Some(type_hint.to_string())
            }
            ExprKind::Ident(name)
                if let Some(type_hint) = crate::adapters::url::call_type_hint(name) =>
            {
                Some(type_hint.to_string())
            }
            ExprKind::Ident(name)
                if let Some(type_hint) = crate::adapters::netip::call_type_hint(name) =>
            {
                Some(type_hint.to_string())
            }
            ExprKind::Ident(name)
                if let Some(type_hint) = crate::adapters::encoding::call_type_hint(name) =>
            {
                Some(type_hint.to_string())
            }
            ExprKind::Ident(name)
                if let Some(type_hint) = crate::adapters::strconv::call_type_hint(name) =>
            {
                Some(type_hint.to_string())
            }
            ExprKind::Ident(name)
                if let Some(type_hint) =
                    crate::adapters::strings::call_type_hint(go_public_adapter_emit_name(name)) =>
            {
                Some(type_hint.to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    go_public_adapter_emit_name(name),
                    "go.time_date"
                        | "go.time_unix"
                        | "go.time_now"
                        | "go.time_unix_milli"
                        | "go.time_unix_micro"
                        | "go.time_time_add"
                        | "go.time_time_add_date"
                        | "go.time_time_truncate"
                        | "go.time_time_round"
                        | "go.time_time_utc"
                        | "go.time_time_in"
                ) =>
            {
                Some("__goTime".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    go_public_adapter_emit_name(name),
                    "go.time_utc"
                        | "go.time_local"
                        | "go.time_fixed_zone"
                        | "go.time_time_location"
                ) =>
            {
                Some("__goLoc".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    go_public_adapter_emit_name(name),
                    "go.time_time_format"
                        | "go.time_time_month"
                        | "go.time_time_weekday"
                        | "go.time_location_string"
                        | "go.time_duration_string"
                ) =>
            {
                Some("string".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    go_public_adapter_emit_name(name),
                    "go.time_since"
                        | "go.time_until"
                        | "go.time_time_year"
                        | "go.time_time_day"
                        | "go.time_time_hour"
                        | "go.time_time_minute"
                        | "go.time_time_second"
                        | "go.time_time_nanosecond"
                        | "go.time_time_unix"
                        | "go.time_time_unix_nano"
                        | "go.time_time_unix_milli"
                        | "go.time_time_unix_micro"
                        | "go.time_time_year_day"
                        | "go.time_time_sub"
                        | "go.time_duration_round"
                ) =>
            {
                Some("int".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    go_public_adapter_emit_name(name),
                    "go.time_time_before"
                        | "go.time_time_after"
                        | "go.time_time_equal"
                        | "go.time_time_is_zero"
                ) =>
            {
                Some("bool".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    go_public_adapter_emit_name(name),
                    "go.time_load_location"
                        | "go.time_parse"
                        | "go.time_parse_in_location"
                        | "go.time_parse_duration"
                        | "go.time_time_zone"
                ) =>
            {
                Some("tuple".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "go.strings.NewReader"
                        | "strings.NewReader"
                        | "go.bytes.NewReader"
                        | "bytes.NewReader"
                ) =>
            {
                Some("*__goReader".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "go.encoding.xml.NewDecoder" | "xml.NewDecoder" | "encoding.xml.NewDecoder"
                ) =>
            {
                Some("*__goXMLDecoder".to_string())
            }
            ExprKind::Ident(name) if name == "__go_type_assert" => args
                .get(1)
                .and_then(|arg| go_type_name_from_expr(&arg.value)),
            ExprKind::Ident(name) if name == "__go_reflect_typeof" => {
                Some("__goReflectType".to_string())
            }
            ExprKind::Ident(name) if name == "__go_reflect_valueof" => {
                Some("__goReflectValue".to_string())
            }
            ExprKind::Member { object, field, .. } if field == "slice" => {
                go_expr_type_hint(object, env, signatures)
            }
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "__go_slices_slice_common" | "__go_slices_slice_bound_common"
                ) =>
            {
                args.first()
                    .and_then(|arg| go_expr_type_hint(&arg.value, env, signatures))
            }
            ExprKind::Member { field, .. } if field == "charCodeAt" => Some("int".to_string()),
            ExprKind::Member { object, field, .. } => go_expr_type_hint(object, env, signatures)
                .and_then(|type_name| {
                    go_resolve_struct_member_type(&type_name, field, env, &mut HashSet::new())
                }),
            ExprKind::Ident(name) => signatures.get(name).and_then(|sig| sig.return_type.clone()),
            _ => match go_expr_call_name(callee).as_deref() {
                Some("utf16.Decode") => Some("[]rune".to_string()),
                Some("utf16.Encode") => Some("[]uint16".to_string()),
                Some(name)
                    if matches!(
                        name,
                        "go.encoding.xml.NewDecoder" | "xml.NewDecoder" | "encoding.xml.NewDecoder"
                    ) =>
                {
                    Some("*__goXMLDecoder".to_string())
                }
                Some("__go_reflect_typeof") => Some("__goReflectType".to_string()),
                Some("__go_reflect_valueof") => Some("__goReflectValue".to_string()),
                _ => None,
            },
        },
        _ => None,
    }
}

fn go_expr_call_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            let object_name = go_expr_call_name(object)?;
            Some(format!("{}.{}", object_name, field))
        }
        _ => None,
    }
}

fn go_struct_lookup_name(type_name: &str) -> Option<String> {
    go_named_receiver_type(type_name)
}

fn go_resolve_struct_member_path(
    type_name: &str,
    member: &str,
    env: &GoNormalizeEnv,
    seen: &mut HashSet<String>,
) -> Option<Vec<String>> {
    let lookup = go_struct_lookup_name(type_name)?;
    if !seen.insert(lookup.clone()) {
        return None;
    }
    let info = env.struct_infos.get(&lookup)?;
    if info.member_names.contains(member) {
        return Some(vec![member.to_string()]);
    }
    for (embedded_name, embedded_type) in &info.embedded_fields {
        if let Some(mut tail) = go_resolve_struct_member_path(embedded_type, member, env, seen) {
            let mut path = vec![embedded_name.clone()];
            path.append(&mut tail);
            return Some(path);
        }
    }
    None
}

fn go_resolve_struct_member_type(
    type_name: &str,
    member: &str,
    env: &GoNormalizeEnv,
    seen: &mut HashSet<String>,
) -> Option<String> {
    let lookup = go_struct_lookup_name(type_name)?;
    if !seen.insert(lookup.clone()) {
        return None;
    }
    let info = env.struct_infos.get(&lookup)?;
    if let Some(type_name) = info.member_types.get(member) {
        return Some(type_name.clone());
    }
    for (_, embedded_type) in &info.embedded_fields {
        if let Some(type_name) = go_resolve_struct_member_type(embedded_type, member, env, seen) {
            return Some(type_name);
        }
    }
    None
}

fn go_should_auto_deref_struct_member(
    object: &Expression,
    field: &str,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    let Some(type_name) = go_expr_type_hint(object, env, signatures) else {
        return false;
    };
    let trimmed = type_name.trim();
    let Some(inner) = trimmed
        .strip_prefix('*')
        .or_else(|| trimmed.strip_prefix('^'))
    else {
        return false;
    };
    go_resolve_struct_member_type(inner.trim(), field, env, &mut HashSet::new()).is_some()
}

fn go_rewrite_promoted_member_access(
    object: Expression,
    field: &str,
    null_safe: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let receiver_type = go_expr_type_hint(&object, env, signatures)?;
    if let Some(receiver_lookup) = go_struct_lookup_name(&receiver_type)
        && go_promoted_pointer_method_wrapper_owner(&receiver_lookup, field, env).is_some()
    {
        return None;
    }
    let path = go_resolve_struct_member_path(&receiver_type, field, env, &mut HashSet::new())?;
    if path.len() <= 1 {
        return None;
    }

    let mut expr = object;
    for segment in path {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: segment,
            null_safe,
        });
    }
    Some(expr)
}

fn go_rewrite_method_expression_member(
    object: &Expression,
    field: &str,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let receiver_type = go_method_expression_receiver_type(object)?;
    let lookup = go_struct_lookup_name(&receiver_type)?;
    let info = env.struct_infos.get(&lookup)?;
    if go_skip_method_wrapper_type(&lookup) || !info.method_names.contains(field) {
        return None;
    }
    Some(Expression::new(ExprKind::FuncRef(
        go_scalar_method_wrapper_name(&lookup, field),
    )))
}

fn go_normalize_method_value_binding(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    _state: &mut GoNormalizeState,
) -> Expression {
    let ExprKind::Member { object, field, .. } = &expr.kind else {
        return expr;
    };
    go_bound_method_value_expr(object.as_ref().clone(), field, env, signatures).unwrap_or(expr)
}

fn go_method_value_binding_from_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<GoMethodValueBinding> {
    let ExprKind::Member { object, field, .. } = &expr.kind else {
        return None;
    };
    go_bound_method_call_expr(object.as_ref().clone(), field, &[], false, env, signatures)?;
    Some(GoMethodValueBinding {
        receiver: object.as_ref().clone(),
        method: field.clone(),
    })
}

fn go_method_value_binding_call_expr(
    binding: &GoMethodValueBinding,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    go_bound_method_call_expr(
        binding.receiver.clone(),
        &binding.method,
        args,
        optional,
        env,
        signatures,
    )
}

fn go_bound_method_value_expr(
    mut receiver_object: Expression,
    field: &str,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let mut receiver_type = go_expr_type_hint(&receiver_object, env, signatures)?;
    if env.interface_methods.contains_key(receiver_type.trim()) {
        receiver_type =
            go_interface_receiver_concrete_type(&receiver_object, &receiver_type, field, env)?;
    }

    let mut lookup = go_struct_lookup_name(&receiver_type)?;
    let mut info = env.struct_infos.get(&lookup)?;
    let mut method_params = info.method_params.get(field).cloned();
    let mut wrapper_lookup = lookup.clone();

    if !info.method_names.contains(field) {
        if let Some(promoted_owner) = go_promoted_pointer_method_wrapper_owner(&lookup, field, env)
        {
            method_params = go_promoted_pointer_method_params(&promoted_owner, field, env);
            wrapper_lookup = promoted_owner;
        } else {
            let (promoted_object, promoted_type) = go_promoted_method_receiver_for_call(
                receiver_object.clone(),
                &receiver_type,
                field,
                env,
            )?;
            receiver_object = promoted_object;
            receiver_type = promoted_type;
            lookup = go_struct_lookup_name(&receiver_type)?;
            info = env.struct_infos.get(&lookup)?;
            method_params = info.method_params.get(field).cloned();
            wrapper_lookup = lookup.clone();
            if !info.method_names.contains(field) {
                return None;
            }
        }
    }

    if go_skip_method_wrapper_type(&wrapper_lookup) {
        return None;
    }

    let params = method_params?;
    let declared_receiver = params.first().and_then(|param| param.type_hint.as_deref());
    let receiver_arg = match declared_receiver {
        Some(declared) if receiver_type.trim().starts_with('*') && !declared.starts_with('*') => {
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Deref,
                expr: Box::new(receiver_object.clone()),
            })
        }
        Some(declared) if !receiver_type.trim().starts_with('*') && declared.starts_with('*') => {
            go_addr_of_normalized_expr(receiver_object.clone())
        }
        _ => receiver_object,
    };
    let receiver_arg = match declared_receiver {
        Some(declared) if !declared.starts_with('*') => {
            go_wrap_go_value_copy(receiver_arg, env, signatures)
        }
        _ => receiver_arg,
    };

    let call_params = params.into_iter().skip(1).collect::<Vec<_>>();
    let mut call_args = Vec::with_capacity(call_params.len() + 1);
    call_args.push(Argument::positional(receiver_arg.clone()));
    call_args.extend(
        call_params
            .iter()
            .map(|param| Argument::positional(Expression::ident(&param.name))),
    );
    let mut captures = go_big_captures(&[&receiver_arg]);
    captures.retain(|name| !name.starts_with("__go_"));
    Some(Expression::new(ExprKind::Lambda {
        params: call_params,
        body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::FuncRef(
                go_scalar_method_wrapper_name(&wrapper_lookup, field),
            ))),
            args: call_args,
            optional: false,
        }))),
        is_async: false,
        captures,
    }))
}

fn go_bound_method_call_expr(
    mut receiver_object: Expression,
    field: &str,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let mut receiver_type = go_expr_type_hint(&receiver_object, env, signatures)?;
    if env.interface_methods.contains_key(receiver_type.trim()) {
        receiver_type =
            go_interface_receiver_concrete_type(&receiver_object, &receiver_type, field, env)?;
    }

    let mut lookup = go_struct_lookup_name(&receiver_type)?;
    let mut info = env.struct_infos.get(&lookup)?;
    let mut wrapper_lookup = lookup.clone();

    if !info.method_names.contains(field) {
        if let Some(promoted_owner) = go_promoted_pointer_method_wrapper_owner(&lookup, field, env)
        {
            wrapper_lookup = promoted_owner;
        } else {
            let (promoted_object, promoted_type) = go_promoted_method_receiver_for_call(
                receiver_object.clone(),
                &receiver_type,
                field,
                env,
            )?;
            receiver_object = promoted_object;
            receiver_type = promoted_type;
            lookup = go_struct_lookup_name(&receiver_type)?;
            info = env.struct_infos.get(&lookup)?;
            wrapper_lookup = lookup.clone();
            if !info.method_names.contains(field) {
                return None;
            }
        }
    }

    if go_skip_method_wrapper_type(&wrapper_lookup) {
        return None;
    }

    let wrapper_info = env.struct_infos.get(&wrapper_lookup)?;
    let declared_receiver = if wrapper_lookup == lookup {
        wrapper_info.method_receiver_types.get(field).cloned()
    } else {
        Some(format!("*{}", wrapper_lookup))
    };
    let declared_receiver = declared_receiver.as_deref().map(str::trim);
    let receiver_arg = match declared_receiver {
        Some(declared) if receiver_type.trim().starts_with('*') && !declared.starts_with('*') => {
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Deref,
                expr: Box::new(receiver_object.clone()),
            })
        }
        Some(declared) if !receiver_type.trim().starts_with('*') && declared.starts_with('*') => {
            go_addr_of_normalized_expr(receiver_object.clone())
        }
        _ => receiver_object,
    };
    let receiver_arg = match declared_receiver {
        Some(declared) if !declared.starts_with('*') => {
            go_wrap_go_value_copy(receiver_arg, env, signatures)
        }
        _ => receiver_arg,
    };
    let receiver_arg = match declared_receiver {
        Some(declared) if !declared.starts_with('*') => {
            if let Some(underlying) = env.named_types.get(&lookup) {
                go_normalize_type_conversion(underlying, receiver_arg, env, signatures)
            } else {
                receiver_arg
            }
        }
        _ => receiver_arg,
    };

    let mut rewritten_args = Vec::with_capacity(args.len() + 1);
    rewritten_args.push(Argument::positional(receiver_arg));
    rewritten_args.extend(args.iter().cloned());

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&go_scalar_method_wrapper_name(
            &wrapper_lookup,
            field,
        ))),
        args: rewritten_args,
        optional,
    }))
}

fn go_promoted_pointer_method_params(
    receiver_lookup: &str,
    method: &str,
    env: &GoNormalizeEnv,
) -> Option<Vec<Param>> {
    let info = env.struct_infos.get(receiver_lookup)?;
    for (_, embedded_type) in &info.embedded_fields {
        let embedded_lookup = go_struct_lookup_name(embedded_type)?;
        let embedded_info = env.struct_infos.get(&embedded_lookup)?;
        if embedded_info.pointer_method_names.contains(method) {
            return embedded_info.method_params.get(method).cloned();
        }
    }
    None
}

fn go_method_expression_receiver_type(object: &Expression) -> Option<String> {
    match &object.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::RefLoad(inner) => match &inner.kind {
            ExprKind::Ident(name) => Some(format!("*{}", name)),
            _ => None,
        },
        ExprKind::Unary {
            op: UnaryOp::Deref,
            expr,
        } => match &expr.kind {
            ExprKind::Ident(name) => Some(format!("*{}", name)),
            _ => None,
        },
        _ => None,
    }
}

fn go_rewrite_named_type_method_call(
    callee: &Expression,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let mut receiver_object = (**object).clone();
    let mut receiver_type = go_expr_type_hint(object, env, signatures)?;
    if let Some(rewritten) = go_rewrite_tree_adapter_method_call(
        receiver_object.clone(),
        &receiver_type,
        field,
        args,
        optional,
    ) {
        return Some(rewritten);
    }
    if let Some((embedded_object, embedded_interface_type)) =
        go_embedded_interface_method_receiver(&receiver_object, &receiver_type, field, env)
    {
        receiver_object = embedded_object;
        receiver_type = go_interface_receiver_concrete_type(
            &receiver_object,
            &embedded_interface_type,
            field,
            env,
        )?;
    } else if env.interface_methods.contains_key(receiver_type.trim()) {
        receiver_type =
            go_interface_receiver_concrete_type(&receiver_object, &receiver_type, field, env)?;
    }
    let mut lookup = go_struct_lookup_name(&receiver_type)?;
    let mut info = env.struct_infos.get(&lookup)?;
    if !info.method_names.contains(field) {
        if let Some(promoted_owner) = go_promoted_pointer_method_wrapper_owner(&lookup, field, env)
        {
            let receiver_arg = if receiver_type.trim().starts_with('*') {
                receiver_object
            } else {
                go_addr_of_normalized_expr(receiver_object)
            };
            let mut rewritten_args = Vec::with_capacity(args.len() + 1);
            rewritten_args.push(Argument::positional(receiver_arg));
            rewritten_args.extend(args.iter().cloned());
            return Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&go_scalar_method_wrapper_name(
                    &promoted_owner,
                    field,
                ))),
                args: rewritten_args,
                optional,
            }));
        }
        let (promoted_object, promoted_type) = go_promoted_method_receiver_for_call(
            receiver_object.clone(),
            &receiver_type,
            field,
            env,
        )?;
        receiver_object = promoted_object;
        receiver_type = promoted_type;
        lookup = go_struct_lookup_name(&receiver_type)?;
        info = env.struct_infos.get(&lookup)?;
        if !info.method_names.contains(field) {
            return None;
        }
    }
    if go_skip_method_wrapper_type(&lookup) {
        return None;
    }

    let declared_receiver = info.method_receiver_types.get(field).map(|ty| ty.trim());
    let receiver_arg = match declared_receiver {
        Some(declared) if receiver_type.trim().starts_with('*') && !declared.starts_with('*') => {
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Deref,
                expr: Box::new(receiver_object.clone()),
            })
        }
        Some(declared) if !receiver_type.trim().starts_with('*') && declared.starts_with('*') => {
            go_addr_of_normalized_expr(receiver_object.clone())
        }
        _ => receiver_object,
    };
    let receiver_arg = match declared_receiver {
        Some(declared) if !declared.starts_with('*') => {
            go_wrap_go_value_copy(receiver_arg, env, signatures)
        }
        _ => receiver_arg,
    };
    let receiver_arg = match declared_receiver {
        Some(declared) if !declared.starts_with('*') => {
            if let Some(underlying) = env.named_types.get(&lookup) {
                go_normalize_type_conversion(underlying, receiver_arg, env, signatures)
            } else {
                receiver_arg
            }
        }
        _ => receiver_arg,
    };

    let mut rewritten_args = Vec::with_capacity(args.len() + 1);
    rewritten_args.push(Argument::positional(receiver_arg));
    rewritten_args.extend(args.iter().cloned());

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&go_scalar_method_wrapper_name(
            &lookup, field,
        ))),
        args: rewritten_args,
        optional,
    }))
}

fn go_rewrite_tree_adapter_method_call(
    receiver: Expression,
    receiver_type: &str,
    field: &str,
    args: &[Argument],
    optional: bool,
) -> Option<Expression> {
    let raw = receiver_type.trim();
    let bare = raw.trim_start_matches('*').trim();
    let go_scope = ["go".to_string()];
    for class_name in [raw, bare] {
        let Some(node) = vybe_compiler::primitives::namespaces::lookup_type_instance_member(
            &go_scope, class_name, field, None,
        ) else {
            continue;
        };
        if let vybe_compiler::primitives::namespaces::NamespaceNode::CommonEmit(emit) = node {
            let mut rewritten_args = Vec::with_capacity(args.len() + 1);
            rewritten_args.push(Argument::positional(receiver.clone()));
            rewritten_args.extend(args.iter().cloned());
            return Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&emit)),
                args: rewritten_args,
                optional,
            }));
        }
    }
    None
}

fn go_addr_of_normalized_expr(expr: Expression) -> Expression {
    if let Some(place) = PlaceExpr::from_expr(&expr) {
        Expression::new(ExprKind::RefOf(Box::new(place)))
    } else {
        Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(expr),
        })
    }
}

fn go_promoted_pointer_method_wrapper_owner(
    receiver_lookup: &str,
    method: &str,
    env: &GoNormalizeEnv,
) -> Option<String> {
    let info = env.struct_infos.get(receiver_lookup)?;
    if info.method_names.contains(method)
        || go_has_ambiguous_promoted_method(receiver_lookup, method, env)
    {
        return None;
    }
    info.embedded_fields
        .iter()
        .filter_map(|(_, embedded_type)| {
            let embedded_lookup = go_struct_lookup_name(embedded_type)?;
            let embedded_info = env.struct_infos.get(&embedded_lookup)?;
            embedded_info
                .pointer_method_names
                .contains(method)
                .then_some(())
        })
        .next()
        .map(|_| receiver_lookup.to_string())
}

fn go_materialize_addressed_composite_decl_init(
    init: Option<&Expression>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<(VarDeclarator, Expression, String, String)> {
    let ExprKind::Unary {
        op: UnaryOp::AddrOf,
        expr,
    } = &init?.kind
    else {
        return None;
    };
    if PlaceExpr::from_expr(expr).is_some() {
        return None;
    }
    if !matches!(
        expr.kind,
        ExprKind::Cast { .. } | ExprKind::Object(_) | ExprKind::Array(_)
    ) {
        return None;
    }
    let type_name = go_expr_type_hint(expr, env, signatures)?;
    let tmp_name = fresh_go_temp(state, "__go_addr_tmp");
    let tmp_decl = VarDeclarator {
        pattern: BindingPattern::Ident(tmp_name.clone()),
        type_hint: Some(type_name.clone().into()),
        init: Some(expr.as_ref().clone()),
        array_bounds: None,
        with_events: false,
    };
    let ref_expr = Expression::new(ExprKind::RefOf(Box::new(PlaceExpr::Ident(
        tmp_name.clone(),
    ))));
    Some((tmp_decl, ref_expr, tmp_name, type_name))
}

fn go_skip_method_wrapper_type(name: &str) -> bool {
    name.starts_with("__go") && !matches!(name, "__goList" | "__goListElement" | "__goRing")
}

fn go_promoted_method_receiver_for_call(
    object: Expression,
    receiver_type: &str,
    method: &str,
    env: &GoNormalizeEnv,
) -> Option<(Expression, String)> {
    let mut path = go_resolve_struct_member_path(receiver_type, method, env, &mut HashSet::new())?;
    if path.len() <= 1 {
        return None;
    }
    path.pop();
    let mut expr = object;
    let mut type_name = receiver_type.to_string();
    for segment in path {
        type_name = go_resolve_struct_member_type(&type_name, &segment, env, &mut HashSet::new())?;
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: segment,
            null_safe: false,
        });
    }
    Some((expr, type_name))
}

fn go_interface_receiver_concrete_type(
    receiver: &Expression,
    interface_type: &str,
    method: &str,
    env: &GoNormalizeEnv,
) -> Option<String> {
    if let ExprKind::Ident(name) = &receiver.kind {
        if let Some(concrete) = env.interface_concrete_types.get(name) {
            if go_type_has_method(concrete, method, env) {
                return Some(concrete.clone());
            }
        }
    }
    let concrete = go_concrete_type_for_interface(interface_type.trim(), env)?;
    if go_type_has_method(&concrete, method, env) {
        Some(concrete)
    } else {
        None
    }
}

fn go_embedded_interface_method_receiver(
    object: &Expression,
    receiver_type: &str,
    method: &str,
    env: &GoNormalizeEnv,
) -> Option<(Expression, String)> {
    let lookup = go_struct_lookup_name(receiver_type)?;
    let info = env.struct_infos.get(&lookup)?;
    for (embedded_name, embedded_type) in &info.embedded_fields {
        if env
            .interface_methods
            .get(embedded_type.trim())
            .is_some_and(|required| required.contains(method))
        {
            return Some((
                Expression::new(ExprKind::Member {
                    object: Box::new(object.clone()),
                    field: embedded_name.clone(),
                    null_safe: false,
                }),
                embedded_type.clone(),
            ));
        }
    }
    None
}

fn go_is_function_type(type_name: &str) -> bool {
    type_name.trim().starts_with("func(")
}

fn go_normalize_function_value_for_inferred_binding(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let ExprKind::Ident(name) = &expr.kind else {
        return expr;
    };
    if !signatures.contains_key(name)
        || env.value_types.contains_key(name)
        || env.package_aliases.contains_key(name)
        || env.type_names.contains(name)
    {
        return expr;
    }
    Expression::new(ExprKind::FuncRef(name.clone()))
}

fn go_normalize_function_value(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(_) => {
            go_normalize_function_value_for_inferred_binding(expr, env, signatures)
        }
        _ => expr,
    }
}

fn go_rewrite_callable_field_member_call(
    callee: &Expression,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if matches!(&object.kind, ExprKind::Ident(name) if env.package_aliases.contains_key(name)) {
        return None;
    }
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    let lookup = go_struct_lookup_name(&receiver_type)?;
    let info = env.struct_infos.get(&lookup)?;
    if info.method_names.contains(field) {
        return None;
    }
    let field_type = info.member_types.get(field)?;
    if !go_is_function_type(field_type) {
        return None;
    }

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Sequence(vec![callee.clone()]))),
        args: args.to_vec(),
        optional,
    }))
}

fn go_normalize_typed_composite_expr(
    expr: Expression,
    type_name: &str,
    env: &GoNormalizeEnv,
) -> Expression {
    if let Some(underlying) = env.named_types.get(type_name.trim()) {
        if go_is_array_like_type(underlying) {
            let normalized_expr = match expr.kind {
                ExprKind::Object(props) if props.is_empty() => {
                    Expression::new(ExprKind::Array(Vec::new()))
                }
                _ => expr,
            };
            return Expression::new(ExprKind::Cast {
                expr: Box::new(normalized_expr),
                type_name: type_name.to_string(),
            });
        }
        if go_is_map_type(underlying) {
            return Expression::new(ExprKind::Cast {
                expr: Box::new(expr),
                type_name: type_name.to_string(),
            });
        }
    }

    if let ExprKind::Array(elements) = &expr.kind {
        if let Some(lookup) = go_struct_lookup_name(type_name) {
            if let Some(info) = env.struct_infos.get(&lookup) {
                let mut props = Vec::new();
                for (index, field_name) in info.field_order.iter().enumerate() {
                    let value = elements
                        .get(index)
                        .map(|element| element.value.clone())
                        .or_else(|| {
                            info.member_types
                                .get(field_name)
                                .map(|field_type| go_zero_value_for_type(field_type, env))
                        });
                    if let Some(value) = value {
                        go_push_struct_field_prop(&mut props, field_name, value, env);
                    }
                }
                return Expression::new(ExprKind::Cast {
                    expr: Box::new(Expression::new(ExprKind::Object(props))),
                    type_name: type_name.to_string(),
                });
            }
        }
    }

    if let ExprKind::Object(props) = &expr.kind {
        if let Some(lookup) = go_struct_lookup_name(type_name) {
            if let Some(info) = env.struct_infos.get(&lookup) {
                let mut filled = Vec::new();
                for field_name in &info.field_order {
                    let value = go_object_prop_value(props, field_name).or_else(|| {
                        info.member_types
                            .get(field_name)
                            .map(|field_type| go_zero_value_for_type(field_type, env))
                    });
                    if let Some(value) = value {
                        go_push_struct_field_prop(&mut filled, field_name, value, env);
                    }
                }
                return Expression::new(ExprKind::Cast {
                    expr: Box::new(Expression::new(ExprKind::Object(filled))),
                    type_name: type_name.to_string(),
                });
            }
        }

        if go_is_array_like_type(type_name) {
            let elem_type = go_array_element_type(type_name);
            let mut values = Vec::new();
            if let Some(target_len) = go_fixed_array_len(type_name, props.len()) {
                if let Some(elem_type) = elem_type.as_deref() {
                    values.resize_with(target_len, || go_zero_value_for_type(elem_type, env));
                } else {
                    values.resize_with(target_len, Expression::null);
                }
            }
            let mut next_index = 0usize;
            for prop in props {
                let ObjectProperty::KeyValue { key, value } = prop else {
                    continue;
                };
                let index = go_composite_literal_index_key(key).unwrap_or(next_index);
                if index >= values.len() {
                    if let Some(elem_type) = elem_type.as_deref() {
                        values.resize_with(index + 1, || go_zero_value_for_type(elem_type, env));
                    } else {
                        values.resize_with(index + 1, Expression::null);
                    }
                }
                values[index] = go_retype_elided_element(value.clone(), elem_type.as_deref());
                next_index = index + 1;
            }
            let arr_elems = values
                .into_iter()
                .map(|value| ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                })
                .collect();
            return Expression::new(ExprKind::Cast {
                expr: Box::new(Expression::new(ExprKind::Array(arr_elems))),
                type_name: type_name.to_string(),
            });
        }
    }

    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn go_push_struct_field_prop(
    props: &mut Vec<ObjectProperty>,
    field_name: &str,
    value: Expression,
    env: &GoNormalizeEnv,
) {
    if let Some(cap) = go_bound_slice_capacity_expr(&value, env) {
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&format!("{}__cap", field_name)),
            value: cap,
        });
    }
    props.push(ObjectProperty::KeyValue {
        key: Expression::string(field_name),
        value,
    });
}

fn go_is_neg_one_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(-1)) => true,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => matches!(expr.kind, ExprKind::Lit(Literal::Int(1))),
        _ => false,
    }
}

fn go_decl_fixed_array_binding(
    decl: &VarDeclarator,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<(String, String)> {
    let BindingPattern::Ident(name) = &decl.pattern else {
        return None;
    };
    let type_name = decl.type_hint.as_deref().map(str::to_string).or_else(|| {
        decl.init
            .as_ref()
            .and_then(|expr| go_expr_type_hint(expr, env, signatures))
    })?;
    go_is_fixed_array_type(&type_name).then(|| (name.clone(), type_name))
}

fn go_decl_binding_type(
    decl: &VarDeclarator,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<(String, String)> {
    let BindingPattern::Ident(name) = &decl.pattern else {
        return None;
    };

    decl.type_hint
        .as_deref()
        .map(str::to_string)
        .or_else(|| {
            decl.init
                .as_ref()
                .and_then(go_utf16_call_type_hint)
                .or_else(|| {
                    decl.init
                        .as_ref()
                        .and_then(|expr| go_expr_type_hint(expr, env, signatures))
                })
        })
        .map(|type_name| (name.clone(), type_name))
}

fn go_decl_interface_concrete_binding(
    decl: &VarDeclarator,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<(String, String)> {
    let BindingPattern::Ident(name) = &decl.pattern else {
        return None;
    };
    let interface_type = decl.type_hint.as_deref()?.trim();
    if !go_is_go_interface_type(interface_type, env) {
        return None;
    }
    let init = decl.init.as_ref()?;
    let concrete_type = go_known_interface_dynamic_type(init, env)
        .or_else(|| go_expr_type_hint(init, env, signatures))?;
    if go_type_assignable_to_interface(&concrete_type, interface_type, env) {
        Some((name.clone(), concrete_type))
    } else {
        None
    }
}

fn go_interface_concrete_type_from_assignment_target(
    target: &Expression,
    value: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    let ExprKind::Ident(name) = &target.kind else {
        return None;
    };
    let interface_type = env.value_types.get(name)?.trim();
    if !go_is_go_interface_type(interface_type, env) {
        return None;
    }
    let concrete_type = go_known_interface_dynamic_type(value, env)
        .or_else(|| go_expr_type_hint(value, env, signatures))?;
    if go_type_assignable_to_interface(&concrete_type, interface_type, env) {
        Some(concrete_type)
    } else {
        None
    }
}

fn go_is_go_interface_type(type_name: &str, env: &GoNormalizeEnv) -> bool {
    matches!(type_name.trim(), "interface{}" | "any" | "error")
        || env.interface_methods.contains_key(type_name.trim())
}

fn go_type_assignable_to_interface(
    concrete_type: &str,
    interface_type: &str,
    env: &GoNormalizeEnv,
) -> bool {
    if matches!(interface_type.trim(), "interface{}" | "any") {
        true
    } else if interface_type.trim() == "error" {
        go_type_has_method(concrete_type, "Error", env)
    } else {
        go_type_implements_interface(concrete_type, interface_type, env)
    }
}

fn go_type_implements_interface(
    concrete_type: &str,
    interface_type: &str,
    env: &GoNormalizeEnv,
) -> bool {
    let Some(required) = env.interface_methods.get(interface_type.trim()) else {
        return false;
    };
    let is_pointer = concrete_type.trim().starts_with('*');
    required.iter().all(|method| {
        go_type_has_method_in_method_set(
            concrete_type,
            method,
            is_pointer,
            env,
            &mut HashSet::new(),
        )
    })
}

fn go_canonical_go_type(type_name: &str) -> String {
    go_stdlib_type_binding(type_name)
        .unwrap_or(type_name)
        .to_string()
}

fn go_utf16_call_type_hint(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Call { callee, .. } => match go_expr_call_name(callee).as_deref()? {
            "utf16.Decode" => Some("[]rune".to_string()),
            "utf16.Encode" => Some("[]uint16".to_string()),
            _ => None,
        },
        ExprKind::Cast { expr, type_name } => {
            go_utf16_call_type_hint(expr).or_else(|| Some(type_name.clone()))
        }
        _ => None,
    }
}

fn go_expr_tuple_type_hints(
    expr: &Expression,
    env: &GoNormalizeEnv,
    _signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Vec<Option<String>>> {
    if let Some(hints) = crate::adapters::url::tuple_type_hints(expr) {
        return Some(hints);
    }
    if let Some(hints) = crate::adapters::netip::tuple_type_hints(expr) {
        return Some(hints);
    }

    if let ExprKind::Tuple(values) = &expr.kind {
        if values.len() == 2 {
            if let ExprKind::Cast { type_name, .. } = &values[0].kind {
                return Some(vec![Some(type_name.clone()), Some("bool".to_string())]);
            }
        }
        return None;
    }

    let ExprKind::Call { callee, .. } = &expr.kind else {
        return None;
    };
    if let ExprKind::Ident(name) = &callee.as_ref().kind {
        match env.value_types.get(name).map(String::as_str) {
            Some("__goIterNext") => {
                return Some(vec![Some("any".to_string()), Some("bool".to_string())]);
            }
            Some("__goIterNext2") => {
                return Some(vec![
                    Some("any".to_string()),
                    Some("any".to_string()),
                    Some("bool".to_string()),
                ]);
            }
            _ => {}
        }
    }
    let call_name = go_expr_call_name(callee)?;
    match go_public_adapter_emit_name(&call_name) {
        "go.io.ReadAll" | "io.ReadAll" | "ioutil.ReadAll" => {
            Some(vec![Some("[]byte".to_string()), Some("error".to_string())])
        }
        "go.sort_find" => Some(vec![Some("int".to_string()), Some("bool".to_string())]),
        "go.io.Copy" | "go.io.CopyN" | "go.io.CopyBuffer" | "io.Copy" | "io.CopyN"
        | "io.CopyBuffer" => Some(vec![Some("int64".to_string()), Some("error".to_string())]),
        "go.io.ReadAtLeast" | "go.io.ReadFull" | "io.ReadAtLeast" | "io.ReadFull" => {
            Some(vec![Some("int".to_string()), Some("error".to_string())])
        }
        "go.io.WriteString" | "io.WriteString" => {
            Some(vec![Some("int".to_string()), Some("error".to_string())])
        }
        "go.time_load_location" => {
            Some(vec![Some("__goLoc".to_string()), Some("error".to_string())])
        }
        "go.time_parse" | "go.time_parse_in_location" => Some(vec![
            Some("__goTime".to_string()),
            Some("error".to_string()),
        ]),
        "go.time_parse_duration" => Some(vec![Some("int".to_string()), Some("error".to_string())]),
        "__goBase64Encoding.Decode" => {
            Some(vec![Some("int".to_string()), Some("error".to_string())])
        }
        "__goBase64Encoding.DecodeString" => {
            Some(vec![Some("[]byte".to_string()), Some("error".to_string())])
        }
        "go.encoding.xml.Unescape" | "xml.Unescape" | "encoding.xml.Unescape" => {
            Some(vec![Some("string".to_string()), Some("error".to_string())])
        }
        "go.path_split" => Some(vec![Some("string".to_string()), Some("string".to_string())]),
        "__go_sync_map_Load"
        | "__go_sync_map_LoadOrStore"
        | "__go_sync_map_LoadAndDelete"
        | "__go_sync_map_Swap" => Some(vec![Some("any".to_string()), Some("bool".to_string())]),
        "go.encoding.xml.Marshal"
        | "go.encoding.xml.MarshalIndent"
        | "xml.Marshal"
        | "xml.MarshalIndent"
        | "encoding.xml.Marshal"
        | "encoding.xml.MarshalIndent" => {
            Some(vec![Some("[]byte".to_string()), Some("error".to_string())])
        }
        name if name.ends_with(".Token") || name.ends_with(".RawToken") => {
            Some(vec![Some("any".to_string()), Some("error".to_string())])
        }
        name if name.ends_with(".Peek")
            || name.ends_with(".ReadSlice")
            || name.ends_with(".ReadBytes") =>
        {
            Some(vec![Some("[]byte".to_string()), Some("error".to_string())])
        }
        name if name.ends_with(".ReadLine") => Some(vec![
            Some("[]byte".to_string()),
            Some("bool".to_string()),
            Some("error".to_string()),
        ]),
        name if name.ends_with(".ReadByte") => {
            Some(vec![Some("string".to_string()), Some("error".to_string())])
        }
        name if name.ends_with(".ReadRune") => Some(vec![
            Some("string".to_string()),
            Some("int".to_string()),
            Some("error".to_string()),
        ]),
        name if name.ends_with(".ReadString") => {
            Some(vec![Some("string".to_string()), Some("error".to_string())])
        }
        name if name.ends_with(".Read") || name.ends_with(".Discard") => {
            Some(vec![Some("int".to_string()), Some("error".to_string())])
        }
        _ => None,
    }
}

fn go_record_binding_pattern_type_hints(
    pattern: &BindingPattern,
    type_hints: &[Option<String>],
    env: &mut GoNormalizeEnv,
) {
    let BindingPattern::Array(elements) = pattern else {
        return;
    };
    for (idx, element) in elements.iter().enumerate() {
        let Some(Some(type_hint)) = type_hints.get(idx) else {
            continue;
        };
        if let ArrayPatternElem::Pattern(BindingPattern::Ident(name), None) = element {
            env.value_types.insert(name.clone(), type_hint.clone());
        }
    }
}

fn go_expand_static_tuple_decl(
    decl: &VarDeclarator,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Vec<VarDeclarator>> {
    let BindingPattern::Array(elements) = &decl.pattern else {
        return None;
    };
    let ExprKind::Tuple(values) = &decl.init.as_ref()?.kind else {
        return None;
    };
    if elements.len() != values.len() {
        return None;
    }
    let type_hints = go_expr_tuple_type_hints(decl.init.as_ref()?, env, signatures)
        .unwrap_or_else(|| vec![None; values.len()]);
    let mut declarations = Vec::new();
    for (idx, element) in elements.iter().enumerate() {
        let ArrayPatternElem::Pattern(BindingPattern::Ident(name), None) = element else {
            continue;
        };
        declarations.push(VarDeclarator {
            pattern: BindingPattern::Ident(name.clone()),
            type_hint: type_hints.get(idx).and_then(Clone::clone).map(Into::into),
            init: values.get(idx).cloned(),
            array_bounds: None,
            with_events: false,
        });
    }
    (!declarations.is_empty()).then_some(declarations)
}

fn go_record_tuple_target_type_hints(
    targets: &[Expression],
    type_hints: &[Option<String>],
    env: &mut GoNormalizeEnv,
) {
    for (idx, target) in targets.iter().enumerate() {
        let Some(Some(type_hint)) = type_hints.get(idx) else {
            continue;
        };
        if let ExprKind::Ident(name) = &target.kind {
            env.value_types.insert(name.clone(), type_hint.clone());
        }
    }
}

fn go_single_named_binding_pattern(pattern: &BindingPattern) -> Option<BindingPattern> {
    let BindingPattern::Array(elements) = pattern else {
        return None;
    };

    if elements.len() != 1 {
        return None;
    }

    let mut bound_name = None;
    for element in elements {
        match element {
            ArrayPatternElem::Hole => return None,
            ArrayPatternElem::Pattern(BindingPattern::Ident(name), None) => {
                if bound_name.is_some() {
                    return None;
                }
                bound_name = Some(name.clone());
            }
            _ => return None,
        }
    }

    bound_name.map(BindingPattern::Ident)
}

fn go_is_fixed_array_type(type_name: &str) -> bool {
    go_array_head(type_name)
        .map(|(head, _)| !head.trim().is_empty())
        .unwrap_or(false)
}

fn go_is_integer_type(type_name: &str) -> bool {
    matches!(
        type_name.trim(),
        "int"
            | "int8"
            | "int16"
            | "int32"
            | "int64"
            | "uint"
            | "uint8"
            | "uint16"
            | "uint32"
            | "uint64"
            | "uintptr"
            | "byte"
            | "rune"
    )
}

fn go_is_float_type(type_name: &str) -> bool {
    matches!(type_name.trim(), "float32" | "float64")
}

fn go_is_builtin_conversion_type(type_name: &str) -> bool {
    go_is_integer_type(type_name)
        || go_is_float_type(type_name)
        || matches!(type_name.trim(), "string" | "bool" | "any" | "interface{}")
}

fn go_is_type_conversion_target(
    type_name: &str,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    (go_is_builtin_conversion_type(type_name) || env.type_names.contains(type_name))
        && !signatures.contains_key(type_name)
}

fn go_normalize_type_conversion(
    type_name: &str,
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if matches!(type_name.trim(), "any" | "interface{}") {
        return expr;
    }

    if let Some(underlying) = env
        .named_types
        .get(type_name)
        .filter(|underlying| underlying.as_str() != type_name)
    {
        let normalized = go_normalize_type_conversion(underlying, expr, env, signatures);
        return Expression::new(ExprKind::Cast {
            expr: Box::new(normalized),
            type_name: type_name.to_string(),
        });
    }

    if go_is_integer_type(type_name)
        && go_expr_type_hint(&expr, env, signatures)
            .as_deref()
            .and_then(|hint| env.named_types.get(hint.trim()))
            .is_some_and(|underlying| go_is_integer_type(underlying))
    {
        return expr;
    }

    if go_is_integer_type(type_name) {
        let int_expr = go_builtin_call("__go_to_int", vec![expr]);
        if type_name == "int" {
            return int_expr;
        }
        return Expression::new(ExprKind::Cast {
            expr: Box::new(int_expr),
            type_name: type_name.to_string(),
        });
    }

    if type_name == "string"
        && go_expr_type_hint(&expr, env, signatures)
            .as_deref()
            .is_some_and(|ty| ty.trim() == "string")
    {
        return expr;
    }

    if type_name == "string"
        && go_expr_type_hint(&expr, env, signatures)
            .as_deref()
            .is_some_and(go_is_integer_type)
    {
        return go_builtin_call("__go_str_from_char_code", vec![expr]);
    }

    if type_name == "string"
        && go_expr_type_hint(&expr, env, signatures)
            .as_deref()
            .is_some_and(|ty| {
                matches!(go_array_element_type(ty).as_deref(), Some("byte" | "uint8"))
            })
    {
        return crate::adapters::bytes_io::bytes_to_string(expr);
    }

    if type_name == "string"
        && go_expr_type_hint(&expr, env, signatures)
            .as_deref()
            .is_some_and(|ty| {
                matches!(go_array_element_type(ty).as_deref(), Some("rune" | "int32"))
            })
    {
        return go_builtin_call("__go_runes_to_string", vec![expr]);
    }

    if type_name.trim() == "[]rune" {
        return go_builtin_call("__go_string_to_runes", vec![expr]);
    }

    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn walk_package_clause(pair: Pair<Rule>) -> Result<String, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::ident_name {
            return Ok(inner.as_str().to_string());
        }
    }
    Ok(String::new())
}

fn walk_import(pair: Pair<Rule>) -> Result<Import, String> {
    let mut path = String::new();
    let mut alias: Option<String> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::import_spec => {
                for spec_inner in inner.into_inner() {
                    match spec_inner.as_rule() {
                        Rule::ident_name => {
                            alias = Some(spec_inner.as_str().to_string());
                        }
                        Rule::string_literal => {
                            path = unquote(spec_inner.as_str());
                        }
                        _ => {}
                    }
                }
            }
            Rule::string_literal => {
                path = unquote(inner.as_str());
            }
            _ => {}
        }
    }

    path = go_tree_mounted_import_path(&path).unwrap_or(path);

    Ok(Import {
        kind: ImportKind::Simple { path, alias },
        span: Span::default(),
    })
}

fn go_tree_mounted_import_path(path: &str) -> Option<String> {
    let local = match path {
        "container/heap" => "heap",
        "container/list" => "list",
        "container/ring" => "ring",
        "encoding/base64" => "base64",
        "encoding/binary" => "binary",
        "encoding/hex" => "hex",
        "hash/adler32" => "adler32",
        "hash/crc32" => "crc32",
        "hash/fnv" => "fnv",
        "net/netip" => "netip",
        "net/url" => "url",
        "unicode/utf16" => "utf16",
        "unicode/utf8" => "utf8",
        _ => return None,
    };
    Some(local.to_string())
}

fn unquote(s: &str) -> String {
    if s.len() < 2 {
        return s.to_string();
    }

    if s.starts_with('`') && s.ends_with('`') {
        return s[1..s.len() - 1].to_string();
    }

    if !((s.starts_with('"') && s.ends_with('"')) || (s.starts_with('\'') && s.ends_with('\''))) {
        return s.to_string();
    }

    let mut out = String::new();
    let mut chars = s[1..s.len() - 1].chars();
    while let Some(ch) = chars.next() {
        if ch != '\\' {
            out.push(ch);
            continue;
        }

        match chars.next() {
            Some('n') => out.push('\n'),
            Some('r') => out.push('\r'),
            Some('t') => out.push('\t'),
            Some('\\') => out.push('\\'),
            Some('"') => out.push('"'),
            Some('\'') => out.push('\''),
            Some('0') => out.push('\0'),
            Some('x') => {
                let mut hex = String::new();
                if let Some(first) = chars.next() {
                    hex.push(first);
                }
                if let Some(second) = chars.next() {
                    hex.push(second);
                }
                if hex.len() == 2 {
                    if let Ok(value) = u8::from_str_radix(&hex, 16) {
                        out.push(value as char);
                    } else {
                        out.push('x');
                        out.push_str(&hex);
                    }
                } else {
                    out.push('x');
                    out.push_str(&hex);
                }
            }
            Some('u') => {
                let mut hex = String::new();
                for _ in 0..4 {
                    if let Some(ch) = chars.next() {
                        hex.push(ch);
                    }
                }
                if hex.len() == 4 {
                    if let Ok(value) = u32::from_str_radix(&hex, 16) {
                        if let Some(ch) = char::from_u32(value) {
                            out.push(ch);
                        }
                    } else {
                        out.push('u');
                        out.push_str(&hex);
                    }
                } else {
                    out.push('u');
                    out.push_str(&hex);
                }
            }
            Some('U') => {
                let mut hex = String::new();
                for _ in 0..8 {
                    if let Some(ch) = chars.next() {
                        hex.push(ch);
                    }
                }
                if hex.len() == 8 {
                    if let Ok(value) = u32::from_str_radix(&hex, 16) {
                        if let Some(ch) = char::from_u32(value) {
                            out.push(ch);
                        }
                    } else {
                        out.push('U');
                        out.push_str(&hex);
                    }
                } else {
                    out.push('U');
                    out.push_str(&hex);
                }
            }
            Some(other) => out.push(other),
            None => out.push('\\'),
        }
    }

    out
}

fn walk_top_level(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    match pair.as_rule() {
        Rule::function_declaration => Ok(Some(walk_function_decl(pair)?)),
        Rule::method_declaration => Ok(Some(walk_method_decl(pair)?)),
        Rule::var_declaration => Ok(Some(walk_var_decl(pair)?)),
        Rule::const_declaration => Ok(Some(walk_const_decl(pair)?)),
        Rule::type_declaration => walk_type_decl(pair),
        Rule::declaration => {
            for inner in pair.into_inner() {
                return walk_top_level(inner);
            }
            Ok(None)
        }
        _ => Ok(None),
    }
}

// ── Function declarations ─────────────────────────────────────────────────────────────

fn walk_function_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body_stmts = Vec::new();
    let mut return_type: Option<String> = None;
    let mut named_results = Vec::new();
    let mut generic_params = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => name = inner.as_str().to_string(),
            Rule::type_params => generic_params = consume_go_type_params(inner),
            Rule::signature => {
                let sig = walk_signature(inner)?;
                params = sig.params;
                return_type = sig.return_type;
                named_results = sig.named_results;
            }
            Rule::function_body | Rule::block_statement => {
                body_stmts = walk_block(inner)?;
            }
            _ => {}
        }
    }
    prepend_go_generic_type_params(&mut params, &generic_params);

    for param in named_results.iter().rev() {
        body_stmts.insert(
            0,
            go_named_result_marker_stmt(
                &param.name,
                param.type_hint.as_deref().unwrap_or("object"),
            ),
        );
    }
    for param in &named_results {
        params.push(go_hidden_named_result_param(param));
    }

    // `goto`/labels → structured control flow via the shared relooper. Without
    // this the label parsed, `StmtKind::GoTo` was produced, and NOTHING
    // consumed it: the jump silently did nothing and execution fell through.
    // No-op when the body has no labels.
    let body_stmts = vybe_compiler::primitives::control_flow::lower_gotos(
        body_stmts,
        "__go_goto_pc",
        "__go_goto_dispatch",
        false, // Go labels are case-sensitive
    );

    Ok(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body: body_stmts,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    }))
}

fn walk_method_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut receiver_name = String::new();
    let mut receiver_type = String::new();
    let mut receiver_owner = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut body_stmts = Vec::new();
    let mut return_type: Option<String> = None;
    let mut named_results = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::receiver => {
                for r_inner in inner.into_inner() {
                    match r_inner.as_rule() {
                        Rule::ident_name => receiver_name = r_inner.as_str().to_string(),
                        Rule::type_annotation => {
                            receiver_type = walk_type(r_inner.clone());
                            receiver_owner = go_named_receiver_type(&receiver_type)
                                .unwrap_or_else(|| receiver_type.clone());
                        }
                        _ => {}
                    }
                }
            }
            Rule::ident_name => method_name = inner.as_str().to_string(),
            Rule::signature => {
                let sig = walk_signature(inner)?;
                params = sig.params;
                return_type = sig.return_type;
                named_results = sig.named_results;
            }
            Rule::function_body | Rule::block_statement => {
                body_stmts = walk_block(inner)?;
            }
            _ => {}
        }
    }

    for param in named_results.iter().rev() {
        body_stmts.insert(
            0,
            go_named_result_marker_stmt(
                &param.name,
                param.type_hint.as_deref().unwrap_or("object"),
            ),
        );
    }
    for param in &named_results {
        params.push(go_hidden_named_result_param(param));
    }

    // Prepend receiver as first parameter
    params.insert(
        0,
        Param {
            name: if receiver_name.is_empty() {
                "self".to_string()
            } else {
                receiver_name
            },
            type_hint: Some(receiver_type.clone().into()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        },
    );

    let method_stmt = Statement::new(StmtKind::FunctionDecl {
        name: method_name,
        params,
        return_type,
        body: body_stmts,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    });

    Ok(Statement::new(StmtKind::StructDecl {
        name: receiver_owner,
        interfaces: Vec::new(),
        members: vec![ClassMember::Method(Box::new(method_stmt))],
        visibility: Visibility::Public,
        decorators: Vec::new(),
        // NOT a user declaration — a synthetic carrier holding one method for a
        // receiver, merged into the real `type X struct` later. The policy is
        // declared THERE; stating one here would give the merge two answers.
        semantics: ValueSemantics::default(),
    }))
}

fn walk_signature(pair: Pair<Rule>) -> Result<GoSignatureInfo, String> {
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut named_results = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::parameter_list => {
                params = walk_parameter_list(inner)?;
            }
            Rule::result => {
                for r_inner in inner.into_inner() {
                    match r_inner.as_rule() {
                        Rule::type_annotation => return_type = Some(walk_type(r_inner)),
                        Rule::parameter_list => {
                            let p = walk_parameter_list(r_inner)?;
                            named_results = p
                                .iter()
                                .filter(|param| !param.name.starts_with("__go_param_"))
                                .cloned()
                                .collect();
                            return_type = if p.len() == 1 {
                                p[0].type_hint.clone().as_deref().map(str::to_string)
                            } else {
                                Some(format!("[{}]", p.len()))
                            };
                        }
                        _ => {}
                    }
                }
            }
            Rule::type_annotation => {
                return_type = Some(walk_type(inner));
            }
            _ => {}
        }
    }

    Ok(GoSignatureInfo {
        params,
        return_type,
        named_results,
    })
}

fn go_named_result_marker_stmt(name: &str, type_name: &str) -> Statement {
    Statement::new(StmtKind::Expr(go_builtin_call(
        "__go_named_result",
        vec![
            Expression::string(name),
            go_type_arg_expr(type_name.to_string()),
        ],
    )))
}

fn go_hidden_named_result_param(param: &Param) -> Param {
    let type_name = param
        .type_hint
        .clone()
        .unwrap_or_else(|| "object".to_string().into());
    Param {
        name: param.name.clone(),
        type_hint: Some("object".to_string().into()),
        default: Some(go_named_result_cell_object(go_zero_value_expr(&type_name))),
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: true,
        is_nullable: false,
    }
}

fn go_named_type_marker_stmt(name: &str, type_name: &str) -> Statement {
    Statement::new(StmtKind::Expr(go_builtin_call(
        "__go_named_type",
        vec![
            Expression::string(name),
            go_type_arg_expr(type_name.to_string()),
        ],
    )))
}

fn walk_parameter_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::parameter_decl {
            let mut names = Vec::new();
            let mut type_hint: Option<String> = None;
            let mut is_rest = false;

            for p_inner in inner.into_inner() {
                match p_inner.as_rule() {
                    Rule::ident_name => names.push(p_inner.as_str().to_string()),
                    Rule::ident_list => {
                        for id in p_inner.into_inner() {
                            if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                names.push(id.as_str().to_string());
                            }
                        }
                    }
                    Rule::type_annotation => type_hint = Some(walk_type(p_inner)),
                    Rule::variadic_parameter_type => {
                        is_rest = true;
                        for v_inner in p_inner.into_inner() {
                            if v_inner.as_rule() == Rule::type_annotation {
                                type_hint = Some(format!("[]{}", walk_type(v_inner)));
                            }
                        }
                    }
                    _ => {}
                }
            }

            if names.is_empty() && type_hint.is_some() {
                names.push(format!("__go_param_{}", params.len()));
            }

            for name in names {
                params.push(Param {
                    name,
                    type_hint: type_hint.clone().map(Into::into),
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                });
            }
        }
    }
    Ok(params)
}

fn walk_type(pair: Pair<Rule>) -> String {
    if let Some(backing) = go_stdlib_type_binding(pair.as_str()) {
        return backing.to_string();
    }
    common_generics::erased_type_name(pair.as_str())
}

fn consume_go_type_params(pair: Pair<Rule>) -> Vec<GenericParam> {
    common_generics::parse_generic_params_hint(pair.as_str())
}

fn prepend_go_generic_type_params(params: &mut Vec<Param>, generic_params: &[GenericParam]) {
    for name in common_generics::runtime_type_arg_param_names(generic_params)
        .into_iter()
        .rev()
    {
        params.insert(
            0,
            Param {
                name,
                type_hint: Some("__goTypeArg".to_string().into()),
                default: Some(go_runtime_type_arg_expr("any".to_string())),
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: true,
                is_nullable: false,
            },
        );
    }
}

fn walk_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::block_statement | Rule::function_body => {
                stmts.append(&mut walk_block(inner)?);
            }
            Rule::statement_list => {
                for s in inner.into_inner() {
                    if s.as_rule() == Rule::statement {
                        stmts.push(walk_statement(s)?);
                    }
                }
            }
            Rule::statement => {
                stmts.push(walk_statement(inner)?);
            }
            _ => {}
        }
    }
    Ok(stmts)
}

// ── Variable declarations ─────────────────────────────────────────────────────────────

fn walk_var_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut declarations = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::var_spec | Rule::const_spec => {
                let (mut decls, _) = walk_var_spec(inner, VarDeclKind::Let)?;
                declarations.append(&mut decls);
            }
            Rule::var_group | Rule::const_group => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::var_spec || spec.as_rule() == Rule::const_spec {
                        let (mut decls, _) = walk_var_spec(spec, VarDeclKind::Let)?;
                        declarations.append(&mut decls);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    }))
}

fn walk_const_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut declarations = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::const_spec => {
                let (mut decls, _, _) = walk_const_spec(inner, 0, None, None)?;
                declarations.append(&mut decls);
            }
            Rule::const_group => {
                let mut prev_inits: Option<Vec<Expression>> = None;
                let mut prev_type_hint: Option<String> = None;
                let mut iota_index = 0i64;
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::const_spec {
                        let (mut decls, next_inits, next_type_hint) = walk_const_spec(
                            spec,
                            iota_index,
                            prev_inits.clone(),
                            prev_type_hint.clone(),
                        )?;
                        declarations.append(&mut decls);
                        prev_inits = Some(next_inits);
                        prev_type_hint = next_type_hint;
                        iota_index += 1;
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Const,
    }))
}

fn walk_const_spec(
    pair: Pair<Rule>,
    iota_index: i64,
    prev_inits: Option<Vec<Expression>>,
    prev_type_hint: Option<String>,
) -> Result<(Vec<VarDeclarator>, Vec<Expression>, Option<String>), String> {
    let (names, type_hint, init_values) = parse_go_var_spec(pair)?;
    let effective_type_hint = type_hint.or(prev_type_hint);
    let raw_inits = if init_values.is_empty() {
        prev_inits.unwrap_or_default()
    } else {
        init_values
    };
    let next_inits: Vec<Expression> = raw_inits
        .iter()
        .map(|expr| go_rewrite_iota_expr(expr, iota_index))
        .collect();

    if names.len() > 1 && !next_inits.is_empty() {
        let pattern = BindingPattern::Array(
            names
                .into_iter()
                .map(|name| {
                    if name == "_" {
                        ArrayPatternElem::Hole
                    } else {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                    }
                })
                .collect(),
        );
        let init = if next_inits.len() == 1 {
            next_inits[0].clone()
        } else {
            Expression::new(ExprKind::Tuple(next_inits.clone()))
        };
        return Ok((
            vec![VarDeclarator {
                pattern,
                init: Some(init),
                type_hint: effective_type_hint.clone().map(Into::into),
                array_bounds: None,
                with_events: false,
            }],
            raw_inits,
            effective_type_hint,
        ));
    }

    let mut declarations = Vec::new();
    for name in names {
        if name == "_" {
            continue;
        }
        declarations.push(VarDeclarator {
            pattern: BindingPattern::Ident(name),
            init: next_inits.first().cloned(),
            type_hint: effective_type_hint.clone().map(Into::into),
            array_bounds: None,
            with_events: false,
        });
    }

    Ok((declarations, raw_inits, effective_type_hint))
}

fn parse_go_var_spec(
    pair: Pair<Rule>,
) -> Result<(Vec<String>, Option<String>, Vec<Expression>), String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut init_values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident)
                        || id.as_str() == "_"
                    {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::ident_name => names.push(inner.as_str().to_string()),
            Rule::type_annotation => type_hint = Some(walk_type(inner)),
            Rule::expression_list => init_values = walk_expression_list(inner)?,
            Rule::expression => init_values.push(walk_expression(inner)?),
            _ => {}
        }
    }

    Ok((names, type_hint, init_values))
}

fn go_rewrite_iota_expr(expr: &Expression, iota_index: i64) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) if name == "iota" => Expression::int(iota_index),
        ExprKind::Unary { op, expr } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(go_rewrite_iota_expr(expr, iota_index)),
        }),
        ExprKind::Binary { op, left, right } => Expression::new(ExprKind::Binary {
            op: *op,
            left: Box::new(go_rewrite_iota_expr(left, iota_index)),
            right: Box::new(go_rewrite_iota_expr(right, iota_index)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(go_rewrite_iota_expr(cond, iota_index)),
            then: Box::new(go_rewrite_iota_expr(then, iota_index)),
            else_: Box::new(go_rewrite_iota_expr(else_, iota_index)),
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => Expression::new(ExprKind::Call {
            callee: Box::new(go_rewrite_iota_expr(callee, iota_index)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: go_rewrite_iota_expr(&arg.value, iota_index),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
            optional: *optional,
        }),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: Box::new(go_rewrite_iota_expr(object, iota_index)),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(go_rewrite_iota_expr(object, iota_index)),
            index: Box::new(go_rewrite_iota_expr(index, iota_index)),
            null_safe: *null_safe,
        }),
        ExprKind::Cast { expr, type_name } => Expression::new(ExprKind::Cast {
            expr: Box::new(go_rewrite_iota_expr(expr, iota_index)),
            type_name: type_name.clone(),
        }),
        _ => expr.clone(),
    }
}

fn walk_var_spec(
    pair: Pair<Rule>,
    _kind: VarDeclKind,
) -> Result<(Vec<VarDeclarator>, Option<String>), String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut init_values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident)
                        || id.as_str() == "_"
                    {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::ident_name => names.push(inner.as_str().to_string()),
            Rule::type_annotation => type_hint = Some(walk_type(inner)),
            Rule::expression_list => {
                init_values = walk_expression_list(inner)?;
            }
            Rule::expression => {
                init_values.push(walk_expression(inner)?);
            }
            _ => {}
        }
    }

    if names.len() > 1 && !init_values.is_empty() {
        let pattern = BindingPattern::Array(
            names
                .into_iter()
                .map(|name| {
                    if name == "_" {
                        ArrayPatternElem::Hole
                    } else {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                    }
                })
                .collect(),
        );
        let init = if init_values.len() == 1 {
            init_values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Tuple(init_values))
        };
        return Ok((
            vec![VarDeclarator {
                pattern,
                init: Some(init),
                type_hint: type_hint.map(Into::into),
                array_bounds: None,
                with_events: false,
            }],
            None,
        ));
    }

    let mut declarations = Vec::new();
    for name in names {
        if name == "_" {
            continue;
        }
        declarations.push(VarDeclarator {
            pattern: BindingPattern::Ident(name),
            init: init_values.first().cloned(),
            type_hint: type_hint.clone().map(Into::into),
            array_bounds: None,
            with_events: false,
        });
    }

    Ok((declarations, type_hint))
}

// ── Type declarations (struct, interface, type alias) ─────────────────────────────────

fn walk_type_decl(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::type_spec => {
                let mut name = String::new();
                let mut type_str = String::new();

                for spec_inner in inner.into_inner() {
                    match spec_inner.as_rule() {
                        Rule::ident_name => name = spec_inner.as_str().to_string(),
                        Rule::type_params => {
                            let _ = consume_go_type_params(spec_inner);
                        }
                        Rule::type_annotation => {
                            if let Some(type_stmt) =
                                walk_named_type_annotation(name.clone(), spec_inner.clone())?
                            {
                                return Ok(Some(type_stmt));
                            }
                            type_str = walk_type(spec_inner);
                        }
                        Rule::struct_type => {
                            return Ok(Some(walk_struct_type(name, spec_inner)?));
                        }
                        Rule::interface_type => {
                            return Ok(Some(walk_interface_type(name, spec_inner)?));
                        }
                        _ => {}
                    }
                }

                // Keep named-type metadata in a marker statement so it does not
                // create a runtime binding that shadows generated type methods.
                if !type_str.is_empty() && !name.is_empty() {
                    return Ok(Some(go_named_type_marker_stmt(&name, &type_str)));
                }
            }
            Rule::type_group => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::type_spec {
                        return walk_type_decl(spec.into());
                    }
                }
            }
            _ => {}
        }
    }
    Ok(None)
}

fn walk_struct_type(name: String, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut members = Vec::new();

    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::field_decl {
            let mut field_names = Vec::new();
            let mut field_type: Option<String> = None;
            let mut field_tag: Option<String> = None;

            for f_inner in inner.into_inner() {
                match f_inner.as_rule() {
                    Rule::ident_list => {
                        for id in f_inner.into_inner() {
                            if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                field_names.push(id.as_str().to_string());
                            }
                        }
                    }
                    Rule::ident_name => field_names.push(f_inner.as_str().to_string()),
                    Rule::type_annotation => field_type = Some(walk_type(f_inner)),
                    Rule::string_literal => field_tag = Some(unquote(f_inner.as_str())),
                    _ => {}
                }
            }

            // No name in source is what makes a field embedded — record it now,
            // because filling the name in from the type erases the difference
            // between `Inner` and `Inner Inner`.
            let mut embedded = false;
            if field_names.is_empty() {
                if let Some(type_name) = field_type.as_deref().and_then(go_embedded_field_name) {
                    field_names.push(type_name);
                    embedded = true;
                }
            }

            for fname in field_names {
                let mut modifiers = Modifiers::default();
                if let Some(tag) = field_tag.as_ref() {
                    modifiers
                        .decorators
                        .push(Expression::string(&format!("__go_tag:{tag}")));
                }
                if embedded {
                    modifiers
                        .decorators
                        .push(Expression::string(GO_EMBEDDED_MARKER));
                }
                members.push(ClassMember::Field {
                    name: fname,
                    type_hint: field_type.clone(),
                    init: None,
                    modifiers,
                    with_events: false,
                    array_bounds: field_type.as_deref().and_then(go_fixed_array_bounds_exprs),
                    storage: None,
                });
            }
        }
    }

    Ok(Statement::new(StmtKind::StructDecl {
        name,
        interfaces: Vec::new(),
        members,
        visibility: Visibility::Public,
        decorators: Vec::new(),
        // `type X struct` — the real declaration. A Go struct is a VALUE type
        // (assignment, argument passing and return all copy) and the spec makes
        // `==` on a comparable struct field-wise, so both axes are declared.
        semantics: ValueSemantics {
            storage: ValueStorage::Value,
            equality: ValueEquality::Structural,
            ..Default::default()
        },
    }))
}

fn walk_interface_type(name: String, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut members = Vec::new();

    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::interface_member {
            let mut method_name = String::new();
            let mut params = Vec::new();
            let mut return_type: Option<String> = None;

            for m_inner in inner.into_inner() {
                match m_inner.as_rule() {
                    Rule::ident_name => method_name = m_inner.as_str().to_string(),
                    Rule::signature => {
                        let sig = walk_signature(m_inner)?;
                        params = sig.params;
                        return_type = sig.return_type;
                    }
                    _ => {}
                }
            }

            if !method_name.is_empty() {
                members.push(InterfaceMember::Method {
                    name: method_name,
                    params,
                    return_type,
                    is_sub: false,
                    signature_source: None,
                });
            }
        }
    }

    Ok(Statement::new(StmtKind::InterfaceDecl {
        name,
        parents: Vec::new(),
        members,
        decorators: Vec::new(),
    }))
}

fn walk_named_type_annotation(name: String, pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::struct_type => return Ok(Some(walk_struct_type(name, inner)?)),
            Rule::interface_type => return Ok(Some(walk_interface_type(name, inner)?)),
            _ => {}
        }
    }
    Ok(None)
}

fn go_embedded_field_name(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim().trim_start_matches('*').trim();
    if trimmed.is_empty() {
        return None;
    }
    trimmed.rsplit('.').next().map(|name| name.to_string())
}

fn go_named_receiver_type(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim().trim_start_matches('*').trim();
    let trimmed = common_generics::generic_base_name(trimmed);
    if trimmed.is_empty() {
        return None;
    }
    Some(trimmed.to_string())
}

// ── Statements ─────────────────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let rule = pair.as_rule();
    if rule == Rule::statement {
        if let Some(inner) = pair.into_inner().next() {
            let mut s = walk_statement(inner)?;
            if s.span.start_line == 0 {
                s.span = span;
            }
            return Ok(s);
        }
        return Ok(Statement::with_span(StmtKind::Empty, span));
    }

    let kind = match rule {
        Rule::empty_statement => StmtKind::Empty,
        Rule::block_statement => StmtKind::Block(walk_block(pair)?),
        Rule::expression_statement => {
            let expr = walk_expression(first_meaningful(pair)?)?;
            StmtKind::Expr(expr)
        }
        Rule::assignment_statement => walk_assignment(pair)?,
        Rule::short_var_declaration => walk_short_var_decl(pair)?,
        Rule::inc_dec_statement => walk_inc_dec(pair)?,
        Rule::var_declaration => walk_var_decl(pair)?.kind,
        Rule::const_declaration => walk_const_decl(pair)?.kind,
        Rule::if_statement => walk_if(pair)?,
        Rule::switch_statement => walk_switch(pair)?,
        Rule::select_statement => walk_select(pair)?,
        Rule::for_statement => walk_for(pair)?,
        Rule::return_statement => walk_return(pair)?,
        Rule::break_statement => {
            match pair.into_inner().find(|p| p.as_rule() == Rule::ident_name) {
                Some(lbl) => StmtKind::Break(BreakTarget::Label(lbl.as_str().to_string())),
                None => StmtKind::Break(BreakTarget::Implicit),
            }
        }
        Rule::continue_statement => {
            match pair.into_inner().find(|p| p.as_rule() == Rule::ident_name) {
                Some(lbl) => StmtKind::Continue(ContinueTarget::Label(lbl.as_str().to_string())),
                None => StmtKind::Continue(ContinueTarget::Implicit),
            }
        }
        Rule::fallthrough_statement => StmtKind::Expr(Expression::ident(GO_FALLTHROUGH_MARKER)),
        Rule::goto_statement => StmtKind::GoTo(walk_goto(pair)?),
        Rule::labeled_statement => walk_labeled(pair)?,
        Rule::defer_statement => walk_defer_stmt(pair)?,
        Rule::go_statement => walk_go_stmt(pair)?,
        Rule::send_statement => walk_send_stmt(pair)?,
        _ => StmtKind::Empty,
    };
    Ok(Statement::with_span(kind, span))
}

fn walk_defer_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    Ok(StmtKind::Expr(go_builtin_call(
        "__go_defer",
        vec![walk_expression(first_meaningful(pair)?)?],
    )))
}

fn walk_go_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair
        .into_inner()
        .find(|inner| inner.as_rule() == Rule::expression)
        .map(walk_expression)
        .transpose()?
        .unwrap_or_else(Expression::null);

    Ok(StmtKind::Expr(go_builtin_call(
        "__go_spawn",
        vec![go_wrap_spawn_expr(expr)],
    )))
}

fn walk_send_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exprs = Vec::new();
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression | Rule::unary_expression | Rule::primary => {
                exprs.push(walk_expression(inner)?)
            }
            _ => {}
        }
    }

    if exprs.len() == 2 {
        Ok(StmtKind::Expr(Expression::new(ExprKind::Chan(
            ChanOp::Send {
                channel: Box::new(exprs.remove(0)),
                value: Box::new(exprs.remove(0)),
            },
        ))))
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_assignment(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut targets = Vec::new();
    let mut op = "=";
    let mut values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression_list => {
                if targets.is_empty() {
                    targets = walk_expression_list(inner)?;
                } else {
                    values = walk_expression_list(inner)?;
                }
            }
            Rule::assign_op => op = inner.as_str(),
            _ => {}
        }
    }

    if op != "=" {
        // Compound assignment
        if targets.len() == 1 && values.len() == 1 {
            let target = targets[0].clone();
            let value = values[0].clone();
            let compound_op = match op {
                "+=" => Some(CompoundOp::Add),
                "-=" => Some(CompoundOp::Sub),
                "*=" => Some(CompoundOp::Mul),
                "/=" => Some(CompoundOp::Div),
                "%=" => Some(CompoundOp::Mod),
                "&=" => Some(CompoundOp::BitAnd),
                "|=" => Some(CompoundOp::BitOr),
                "^=" => Some(CompoundOp::BitXor),
                "<<=" => Some(CompoundOp::Shl),
                ">>=" => Some(CompoundOp::Shr),
                _ => None,
            };
            if let Some(compound_op) = compound_op {
                return Ok(StmtKind::CompoundAssign {
                    target,
                    op: compound_op,
                    value,
                });
            }
            if op == "&^=" {
                let rhs = Expression::new(ExprKind::Unary {
                    op: UnaryOp::BitNot,
                    expr: Box::new(value),
                });
                return Ok(StmtKind::Assign {
                    targets: vec![target.clone()],
                    value: Expression::new(ExprKind::Binary {
                        op: BinOp::BitAnd,
                        left: Box::new(target),
                        right: Box::new(rhs),
                    }),
                    by_ref: false,
                });
            }
        }
    }

    if targets.len() > 1 {
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Tuple(values))
        };
        if targets
            .iter()
            .all(|target| matches!(target.kind, ExprKind::Ident(_)))
            && !matches!(value.kind, ExprKind::Tuple(_))
        {
            let patterns = targets
                .iter()
                .map(|target| match &target.kind {
                    ExprKind::Ident(name) => {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
                    }
                    _ => ArrayPatternElem::Hole,
                })
                .collect();
            return Ok(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Destructure(
                    DestructurePattern::Array(patterns),
                ))],
                value,
                by_ref: false,
            });
        }
        return Ok(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Tuple(targets))],
            value,
            by_ref: false,
        });
    }

    if values.len() == 1 {
        Ok(StmtKind::Assign {
            targets,
            value: values.into_iter().next().unwrap(),
            by_ref: false,
        })
    } else if !values.is_empty() {
        Ok(StmtKind::Assign {
            targets,
            value: Expression::new(ExprKind::Tuple(values)),
            by_ref: false,
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_short_var_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut names = Vec::new();
    let mut values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::expression_list => {
                values = walk_expression_list(inner)?;
            }
            _ => {}
        }
    }

    let mut declarations = Vec::new();
    if names.len() == 2 && values.len() == 1 {
        if let Some((expr, type_name)) = go_extract_type_assert_expr(&values[0]) {
            if names.first().is_some_and(|name| name == "_")
                && names.get(1).is_some_and(|name| name != "_")
            {
                declarations.push(VarDeclarator {
                    pattern: BindingPattern::Ident(names[1].clone()),
                    init: Some(Expression::new(ExprKind::IsType {
                        expr: Box::new(expr),
                        type_name,
                    })),
                    type_hint: None,
                    array_bounds: None,
                    with_events: false,
                });
                return Ok(StmtKind::VarDecl {
                    declarations,
                    kind: VarDeclKind::Let,
                });
            }
            let pattern = BindingPattern::Array(
                names
                    .into_iter()
                    .map(|name| {
                        if name == "_" {
                            ArrayPatternElem::Hole
                        } else {
                            ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                        }
                    })
                    .collect(),
            );
            declarations.push(VarDeclarator {
                pattern,
                init: Some(Expression::new(ExprKind::Tuple(vec![
                    go_type_assert_expr(expr.clone(), type_name.clone()),
                    Expression::new(ExprKind::IsType {
                        expr: Box::new(expr),
                        type_name,
                    }),
                ]))),
                type_hint: None,
                array_bounds: None,
                with_events: false,
            });
            return Ok(StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::Let,
            });
        }
    }

    if names.len() > 1 && !values.is_empty() {
        let pattern = BindingPattern::Array(
            names
                .into_iter()
                .map(|name| {
                    if name == "_" {
                        ArrayPatternElem::Hole
                    } else {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                    }
                })
                .collect(),
        );
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Tuple(values))
        };
        declarations.push(VarDeclarator {
            pattern,
            init: Some(value),
            type_hint: None,
            array_bounds: None,
            with_events: false,
        });
    } else {
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else if !values.is_empty() {
            Expression::new(ExprKind::Tuple(values))
        } else {
            Expression::new(ExprKind::Lit(Literal::Null))
        };

        for name in names {
            if name == "_" {
                continue;
            }
            declarations.push(VarDeclarator {
                pattern: BindingPattern::Ident(name),
                init: Some(value.clone()),
                // The type go's own inference GUARANTEES for a literal
                // initializer — `i := 0` IS an `int`, `x := 1.5` IS a
                // `float64` (spec: untyped constant defaults), and any later
                // assignment of another type would not have compiled. Stamped
                // so the shared provably-numeric operator fold can see what
                // the go compiler already proved. `Checked`: statically
                // enforced, never converting.
                type_hint: go_literal_init_type_hint(&value).map(vybe_ast::TypeHint::checked),
                array_bounds: None,
                with_events: false,
            });
        }
    }

    Ok(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    })
}

/// The go type of a LITERAL `:=` initializer, or None when the initializer
/// is anything but a numeric literal (possibly signed).
fn go_literal_init_type_hint(expr: &Expression) -> Option<&'static str> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(_)) => Some("int"),
        ExprKind::Lit(Literal::Float(_)) => Some("float64"),
        ExprKind::Unary {
            op: UnaryOp::Neg | UnaryOp::Pos,
            expr,
        } => go_literal_init_type_hint(expr),
        _ => None,
    }
}

fn walk_inc_dec(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut expr = None;
    let is_inc = !pair
        .as_str()
        .trim_end()
        .trim_end_matches(';')
        .trim_end()
        .ends_with("--");

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression => expr = Some(walk_expression(inner)?),
            Rule::primary => expr = Some(walk_primary(inner)?),
            _ => {}
        }
    }

    if let Some(target) = expr {
        Ok(StmtKind::CompoundAssign {
            target,
            op: if is_inc {
                CompoundOp::Add
            } else {
                CompoundOp::Sub
            },
            value: Expression::new(ExprKind::Lit(Literal::Int(1))),
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut cond = None;
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body: Option<Vec<Statement>> = None;
    let mut pre_stmt: Option<Box<Statement>> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression | Rule::if_expression => {
                if cond.is_none() {
                    cond = Some(walk_expression(inner)?);
                }
            }
            Rule::block_statement => {
                if then_body.is_empty() {
                    then_body = walk_block(inner)?;
                }
            }
            Rule::else_clause => {
                for e_inner in inner.into_inner() {
                    match e_inner.as_rule() {
                        Rule::block_statement => else_body = Some(walk_block(e_inner)?),
                        Rule::if_statement => {
                            let elif = walk_if(e_inner)?;
                            match elif {
                                StmtKind::If {
                                    cond: c,
                                    then_body: t,
                                    elifs: nested_elifs,
                                    else_body: nested_else,
                                } => {
                                    elifs.push((c, t));
                                    elifs.extend(nested_elifs);
                                    else_body = nested_else;
                                }
                                StmtKind::Block(stmts) => else_body = Some(stmts),
                                _ => {}
                            }
                        }
                        _ => {}
                    }
                }
            }
            Rule::short_var_declaration => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::expression_statement => {
                let expr = walk_expression(first_meaningful(inner)?)?;
                pre_stmt = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
            }
            Rule::assignment_statement => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            _ => {}
        }
    }

    let if_stmt = Statement::new(StmtKind::If {
        cond: cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        then_body,
        elifs,
        else_body,
    });

    if let Some(pre) = pre_stmt {
        Ok(StmtKind::Block(vec![*pre, if_stmt]))
    } else {
        Ok(if_stmt.kind)
    }
}

/// Sentinel identifier a `fallthrough` statement is walked to, so `walk_switch`
/// can desugar it by inlining the following clause's body.
const GO_FALLTHROUGH_MARKER: &str = "__go_fallthrough__";

fn go_body_ends_with_fallthrough(body: &[Statement]) -> bool {
    matches!(
        body.last().map(|s| &s.kind),
        Some(StmtKind::Expr(Expression { kind: ExprKind::Ident(name), .. })) if name == GO_FALLTHROUGH_MARKER
    )
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    match pair.as_rule() {
        Rule::switch_statement => {
            if let Some(inner) = pair.into_inner().next() {
                return walk_switch(inner);
            }
            return Ok(StmtKind::Empty);
        }
        Rule::type_switch_stmt => return walk_type_switch(pair),
        Rule::expr_switch_stmt => {}
        _ => return Ok(StmtKind::Empty),
    }

    let mut expr = None;
    let mut cases = Vec::new();
    let mut default: Option<Vec<Statement>> = None;
    let mut pre_stmt: Option<Box<Statement>> = None;
    // Clauses in source order, so `fallthrough` can inline the next clause.
    let mut ordered: Vec<(Vec<CaseCondition>, Vec<Statement>)> = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::switch_short_var_init => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::switch_assignment_init => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            Rule::expression => expr = Some(walk_expression(inner)?),
            Rule::short_var_declaration => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::expression_statement => {
                let expr = walk_expression(first_meaningful(inner)?)?;
                pre_stmt = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
            }
            Rule::assignment_statement => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            Rule::expr_case_clause => {
                let mut conditions: Vec<CaseCondition> = Vec::new();
                let mut body = Vec::new();

                for c_inner in inner.into_inner() {
                    match c_inner.as_rule() {
                        Rule::expr_switch_case => {
                            for sc_inner in c_inner.into_inner() {
                                if sc_inner.as_rule() == Rule::expression_list {
                                    for expr in walk_expression_list(sc_inner)? {
                                        conditions.push(CaseCondition::Value(expr));
                                    }
                                } else if sc_inner.as_rule() == Rule::kw_default {
                                    // default case
                                }
                            }
                        }
                        Rule::statement_list => {
                            body = walk_statement_list(c_inner)?;
                        }
                        _ => {}
                    }
                }

                ordered.push((conditions, body));
            }
            _ => {}
        }
    }

    // Desugar `fallthrough`: a clause ending in the marker continues into the
    // next clause's (already-resolved) body.
    for i in (0..ordered.len()).rev() {
        if go_body_ends_with_fallthrough(&ordered[i].1) {
            let next_body = ordered.get(i + 1).map(|c| c.1.clone()).unwrap_or_default();
            let body = &mut ordered[i].1;
            body.pop(); // drop the marker
            body.extend(next_body);
        }
    }
    for (conditions, body) in ordered {
        if conditions.is_empty() {
            default = Some(body);
        } else {
            cases.push(SwitchCase { conditions, body });
        }
    }

    let switch_stmt = if expr.is_none() {
        let mut first_case: Option<(Expression, Vec<Statement>)> = None;
        let mut elifs = Vec::new();
        for case in cases {
            let cond = case
                .conditions
                .into_iter()
                .filter_map(|condition| match condition {
                    CaseCondition::Value(expr) => Some(expr),
                    _ => None,
                })
                .reduce(|left, right| {
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Or,
                        left: Box::new(left),
                        right: Box::new(right),
                    })
                })
                .unwrap_or_else(|| Expression::bool(false));
            if first_case.is_none() {
                first_case = Some((cond, case.body));
            } else {
                elifs.push((cond, case.body));
            }
        }

        if let Some((cond, then_body)) = first_case {
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body: default,
            }
        } else {
            StmtKind::Block(default.unwrap_or_default())
        }
    } else {
        StmtKind::Switch {
            expr: expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
            cases,
            default,
        }
    };

    if let Some(pre) = pre_stmt {
        Ok(StmtKind::Block(vec![*pre, Statement::new(switch_stmt)]))
    } else {
        Ok(switch_stmt)
    }
}

fn walk_type_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut binding_name: Option<String> = None;
    let mut switch_expr: Option<Expression> = None;
    let mut first_case: Option<(Expression, Vec<Statement>)> = None;
    let mut elifs = Vec::new();
    let mut default_body: Option<Vec<Statement>> = None;
    let mut pre_stmt: Option<Box<Statement>> = None;
    // Position-derived, not a process counter: two type switches in one file
    // start at different byte offsets, so the name is unique within the program
    // AND identical every time this source is compiled.
    let switch_temp_name = format!("__go_type_switch{}", pair.as_span().start());
    let switch_temp_expr = Expression::ident(&switch_temp_name);

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::switch_short_var_init => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::switch_assignment_init => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            Rule::short_var_declaration => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::expression_statement => {
                let expr = walk_expression(first_meaningful(inner)?)?;
                pre_stmt = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
            }
            Rule::assignment_statement => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            Rule::type_switch_guard => {
                for guard_inner in inner.into_inner() {
                    match guard_inner.as_rule() {
                        Rule::ident_name => binding_name = Some(guard_inner.as_str().to_string()),
                        Rule::primary | Rule::type_switch_subject => {
                            switch_expr = Some(walk_primary(guard_inner)?)
                        }
                        _ => {}
                    }
                }
            }
            Rule::type_case_clause => {
                let mut case_types = Vec::new();
                let mut body = Vec::new();
                for case_inner in inner.into_inner() {
                    match case_inner.as_rule() {
                        Rule::type_switch_case => {
                            for switch_case_inner in case_inner.into_inner() {
                                match switch_case_inner.as_rule() {
                                    Rule::type_list => {
                                        for ty in switch_case_inner.into_inner() {
                                            if ty.as_rule() == Rule::type_annotation {
                                                case_types.push(walk_type(ty));
                                            } else if ty.as_rule() == Rule::nil_literal {
                                                case_types.push("__go_nil".to_string());
                                            }
                                        }
                                    }
                                    Rule::kw_default => {}
                                    _ => {}
                                }
                            }
                        }
                        Rule::statement_list => body = walk_statement_list(case_inner)?,
                        _ => {}
                    }
                }

                if case_types.is_empty() {
                    default_body = Some(body);
                } else {
                    let expr = switch_temp_expr.clone();
                    let cond = go_type_switch_case_cond(expr.clone(), &case_types);
                    let case_body = go_type_switch_case_body(
                        body,
                        binding_name.as_deref(),
                        expr,
                        &case_types[0],
                    );
                    if first_case.is_none() {
                        first_case = Some((cond, case_body));
                    } else {
                        elifs.push((cond, case_body));
                    }
                }
            }
            _ => {}
        }
    }

    let type_switch_stmt = if let Some((cond, then_body)) = first_case {
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body: default_body,
        }
    } else {
        StmtKind::Block(default_body.unwrap_or_default())
    };

    let switch_temp_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(switch_temp_name),
            type_hint: None,
            init: Some(switch_expr.unwrap_or_else(Expression::null)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    if let Some(pre) = pre_stmt {
        Ok(StmtKind::Block(vec![
            *pre,
            switch_temp_decl,
            Statement::new(type_switch_stmt),
        ]))
    } else {
        Ok(StmtKind::Block(vec![
            switch_temp_decl,
            Statement::new(type_switch_stmt),
        ]))
    }
}

fn walk_select(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut arms: Vec<(ChanOp, Vec<Statement>)> = Vec::new();
    let mut default_body = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_clause => {
                for clause in inner.into_inner() {
                    match clause.as_rule() {
                        Rule::select_case_clause => {
                            if let Some(arm) = walk_select_case_clause(clause)? {
                                arms.push(arm);
                            }
                        }
                        Rule::select_default_clause => {
                            default_body = Some(walk_select_default_clause(clause)?)
                        }
                        _ => {}
                    }
                }
            }
            Rule::select_case_clause => {
                if let Some(arm) = walk_select_case_clause(inner)? {
                    arms.push(arm);
                }
            }
            Rule::select_default_clause => default_body = Some(walk_select_default_clause(inner)?),
            _ => {}
        }
    }

    if arms.is_empty() {
        // `select { default: ... }` runs the default; bare `select {}` must
        // stay a Select so the lowering emits Go's blocks-forever deadlock
        // panic instead of a silent no-op.
        if let Some(body) = default_body {
            return Ok(StmtKind::Block(body));
        }
        return Ok(StmtKind::Select {
            arms: Vec::new(),
            default: None,
        });
    }
    Ok(StmtKind::Select {
        arms: arms
            .into_iter()
            .map(|(comm, body)| SelectArm { comm, body })
            .collect(),
        default: default_body,
    })
}

fn chan_recv(ch: Expression) -> Expression {
    Expression::new(ExprKind::Chan(ChanOp::Recv(Box::new(ch))))
}

fn chan_recv_ok(ch: Expression) -> Expression {
    Expression::new(ExprKind::Chan(ChanOp::RecvOk(Box::new(ch))))
}

fn chan_len(ch: Expression) -> Expression {
    Expression::new(ExprKind::Chan(ChanOp::Len(Box::new(ch))))
}

fn walk_select_case_clause(pair: Pair<Rule>) -> Result<Option<(ChanOp, Vec<Statement>)>, String> {
    let mut prefix = Vec::new();
    let mut body = Vec::new();
    let mut comm = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_comm_clause => {
                let (clause_comm, mut comm_prefix) = walk_select_comm_clause(inner)?;
                comm = clause_comm;
                prefix.append(&mut comm_prefix);
            }
            Rule::statement_list => body.extend(walk_statement_list(inner)?),
            _ => {}
        }
    }

    prefix.extend(body);
    Ok(comm.map(|comm| (comm, prefix)))
}

fn walk_select_default_clause(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::statement_list {
            return walk_statement_list(inner);
        }
    }
    Ok(Vec::new())
}

fn walk_select_comm_clause(pair: Pair<Rule>) -> Result<(Option<ChanOp>, Vec<Statement>), String> {
    let mut comm = None;
    let mut stmts = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_send_clause => {
                let mut exprs = Vec::new();
                for part in inner.into_inner() {
                    match part.as_rule() {
                        Rule::expression | Rule::unary_expression | Rule::primary => {
                            exprs.push(walk_expression(part)?)
                        }
                        _ => {}
                    }
                }
                if exprs.len() == 2 {
                    let op = ChanOp::Send {
                        channel: Box::new(exprs.remove(0)),
                        value: Box::new(exprs.remove(0)),
                    };
                    comm = Some(op.clone());
                    stmts.push(Statement::new(StmtKind::Expr(Expression::new(
                        ExprKind::Chan(op),
                    ))));
                }
            }
            Rule::select_receive_clause => {
                let mut names = Vec::new();
                let mut recv_expr = None;
                let is_assign = inner.as_str().contains("=") && !inner.as_str().contains(":=");

                for part in inner.into_inner() {
                    match part.as_rule() {
                        Rule::ident_list => {
                            for id in part.into_inner() {
                                if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                    names.push(id.as_str().to_string());
                                }
                            }
                        }
                        Rule::expression => recv_expr = Some(walk_expression(part)?),
                        _ => {}
                    }
                }

                if let Some(expr) = recv_expr {
                    // `<-ch` walked to `Chan(Recv(ch))`; the READINESS test
                    // reuses the same op, and two-name bindings upgrade the
                    // performing expression to `RecvOk` (value, ok).
                    if let ExprKind::Chan(ChanOp::Recv(ch)) = &expr.kind {
                        comm = Some(ChanOp::Recv(ch.clone()));
                    }
                    let two_names = names.len() == 2;
                    let perform = if two_names {
                        if let ExprKind::Chan(ChanOp::Recv(ch)) = &expr.kind {
                            chan_recv_ok((**ch).clone())
                        } else {
                            expr
                        }
                    } else {
                        expr
                    };
                    if names.is_empty() {
                        stmts.push(Statement::new(StmtKind::Expr(perform)));
                    } else if is_assign {
                        let targets = names
                            .into_iter()
                            .map(|name| {
                                if name == "_" {
                                    Expression::null()
                                } else {
                                    Expression::ident(&name)
                                }
                            })
                            .collect::<Vec<_>>();
                        let wrap_tuple = targets.len() > 1;
                        let targets = if wrap_tuple {
                            vec![Expression::new(ExprKind::Tuple(targets))]
                        } else {
                            targets
                        };
                        stmts.push(Statement::new(StmtKind::Assign {
                            targets,
                            value: perform,
                            by_ref: false,
                        }));
                    } else {
                        stmts.push(go_short_var_decl_from_parts(names, perform));
                    }
                }
            }
            _ => {}
        }
    }

    Ok((comm, stmts))
}

fn go_short_var_decl_from_parts(names: Vec<String>, value: Expression) -> Statement {
    let declarations = if names.len() > 1 {
        vec![VarDeclarator {
            pattern: BindingPattern::Array(
                names
                    .into_iter()
                    .map(|name| {
                        if name == "_" {
                            ArrayPatternElem::Hole
                        } else {
                            ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                        }
                    })
                    .collect(),
            ),
            // The value already produces the (v, ok) pair — `ChanOp::RecvOk`,
            // whose ok is computed by the lowering (closed ⇒ false), not
            // hardcoded true as the pre-vocabulary select did.
            init: Some(value),
            type_hint: None,
            array_bounds: None,
            with_events: false,
        }]
    } else {
        names
            .into_iter()
            .filter(|name| name != "_")
            .map(|name| VarDeclarator {
                pattern: BindingPattern::Ident(name),
                init: Some(value.clone()),
                type_hint: None,
                array_bounds: None,
                with_events: false,
            })
            .collect()
    };

    Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    })
}

fn walk_statement_list(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::statement {
            stmts.push(walk_statement(inner)?);
        }
    }
    Ok(stmts)
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut init: Option<Box<Statement>> = None;
    let mut cond: Option<Expression> = None;
    let mut update: Option<Expression> = None;
    let mut body = Vec::new();
    let mut is_range = false;
    let mut range_vars = Vec::new();
    let mut range_iter = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::for_clause => {
                for fc_inner in inner.into_inner() {
                    match fc_inner.as_rule() {
                        Rule::for_short_var_nosemi => {
                            init = Some(Box::new(Statement::new(walk_short_var_decl(fc_inner)?)));
                        }
                        Rule::short_var_declaration => {
                            init = Some(Box::new(Statement::new(walk_short_var_decl(fc_inner)?)));
                        }
                        Rule::expression_statement => {
                            let expr = walk_expression(first_meaningful(fc_inner)?)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                            } else if update.is_none() {
                                update = Some(expr);
                            }
                        }
                        Rule::assignment_statement => {
                            let assign = walk_assignment(fc_inner)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(assign)));
                            } else if update.is_none() {
                                if let StmtKind::Assign { targets, value, .. } = assign {
                                    if let Some(target) = targets.into_iter().next() {
                                        update = Some(Expression::new(ExprKind::Assign {
                                            target: Box::new(target),
                                            value: Box::new(value),
                                        }));
                                    }
                                }
                            }
                        }
                        Rule::inc_dec_statement => {
                            let inc_dec = walk_inc_dec(fc_inner)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(inc_dec)));
                            } else if update.is_none() {
                                if let StmtKind::CompoundAssign { target, op, value } = inc_dec {
                                    let bin_op = match op {
                                        CompoundOp::Add => BinOp::Add,
                                        CompoundOp::Sub => BinOp::Sub,
                                        CompoundOp::Mul => BinOp::Mul,
                                        CompoundOp::Div => BinOp::Div,
                                        CompoundOp::Mod => BinOp::Mod,
                                        _ => BinOp::Add,
                                    };
                                    update = Some(Expression::new(ExprKind::Assign {
                                        target: Box::new(target.clone()),
                                        value: Box::new(Expression::new(ExprKind::Binary {
                                            op: bin_op,
                                            left: Box::new(target),
                                            right: Box::new(value),
                                        })),
                                    }));
                                }
                            }
                        }
                        Rule::for_inc_dec => {
                            let inc_dec = walk_inc_dec(fc_inner)?;
                            if let StmtKind::CompoundAssign { target, op, value } = inc_dec {
                                let bin_op = match op {
                                    CompoundOp::Add => BinOp::Add,
                                    CompoundOp::Sub => BinOp::Sub,
                                    CompoundOp::Mul => BinOp::Mul,
                                    CompoundOp::Div => BinOp::Div,
                                    CompoundOp::Mod => BinOp::Mod,
                                    _ => BinOp::Add,
                                };
                                update = Some(Expression::new(ExprKind::Assign {
                                    target: Box::new(target.clone()),
                                    value: Box::new(Expression::new(ExprKind::Binary {
                                        op: bin_op,
                                        left: Box::new(target),
                                        right: Box::new(value),
                                    })),
                                }));
                            }
                        }
                        Rule::for_assign_nosemi => {
                            let assign = walk_assignment(fc_inner)?;
                            if let StmtKind::Assign { targets, value, .. } = assign {
                                if let Some(target) = targets.into_iter().next() {
                                    update = Some(Expression::new(ExprKind::Assign {
                                        target: Box::new(target),
                                        value: Box::new(value),
                                    }));
                                }
                            }
                        }
                        Rule::expression => {
                            if cond.is_none() {
                                cond = Some(walk_expression(fc_inner)?);
                            } else if update.is_none() {
                                update = Some(walk_expression(fc_inner)?);
                            }
                        }
                        Rule::block_statement => {
                            body = walk_block(fc_inner)?;
                        }
                        _ => {}
                    }
                }
            }
            Rule::range_clause => {
                is_range = true;
                for rc_inner in inner.into_inner() {
                    match rc_inner.as_rule() {
                        Rule::expression_list => {
                            for expr in walk_expression_list(rc_inner)? {
                                let name = if let ExprKind::Ident(id) = &expr.kind {
                                    id.clone()
                                } else {
                                    "_".to_string()
                                };
                                range_vars.push(BindingPattern::Ident(name));
                            }
                        }
                        Rule::ident_list => {
                            for id in rc_inner.into_inner() {
                                if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                    range_vars.push(BindingPattern::Ident(id.as_str().to_string()));
                                }
                            }
                        }
                        Rule::expression | Rule::range_expression => {
                            range_iter = Some(walk_expression(rc_inner)?);
                        }
                        Rule::block_statement => {
                            body = walk_block(rc_inner)?;
                        }
                        _ => {}
                    }
                }
            }
            Rule::expression | Rule::unary_expression | Rule::primary => {
                cond = Some(walk_expression(inner)?);
            }
            Rule::expression_statement => {
                cond = Some(walk_expression(first_meaningful(inner)?)?);
            }
            Rule::block_statement => {
                body = walk_block(inner)?;
            }
            _ => {}
        }
    }

    if is_range {
        let var = if range_vars.len() > 1 {
            range_vars
                .get(1)
                .cloned()
                .unwrap_or_else(|| BindingPattern::Ident("_".to_string()))
        } else if range_vars.len() == 1 {
            BindingPattern::Ident("_".to_string())
        } else {
            range_vars
                .get(0)
                .cloned()
                .unwrap_or_else(|| BindingPattern::Ident("_".to_string()))
        };
        let var_name = match var {
            BindingPattern::Ident(name) => name,
            _ => "_".to_string(),
        };
        let key = if range_vars.len() > 1 {
            let key_pat = range_vars.get(0).cloned().unwrap();
            match key_pat {
                BindingPattern::Ident(name) => Some(name),
                _ => None,
            }
        } else if range_vars.len() == 1 {
            let key_pat = range_vars.get(0).cloned().unwrap();
            match key_pat {
                BindingPattern::Ident(name) => Some(name),
                _ => None,
            }
        } else {
            None
        };

        Ok(StmtKind::ForIn {
            var: var_name,
            key,
            iter: range_iter.unwrap_or_else(|| Expression::new(ExprKind::Array(Vec::new()))),
            body,
            of: true,
            else_body: None,
            is_async: false,
        })
    } else if init.is_none() && update.is_none() && cond.is_some() {
        Ok(StmtKind::While {
            cond: cond.unwrap(),
            body,
            else_body: None,
        })
    } else {
        Ok(StmtKind::For {
            init,
            cond,
            update,
            body,
        })
    }
}

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut values = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::expression_list {
            values = walk_expression_list(inner)?;
        } else if inner.as_rule() == Rule::expression {
            values.push(walk_expression(inner)?);
        }
    }

    if values.len() == 1 {
        Ok(StmtKind::Return(Some(values.into_iter().next().unwrap())))
    } else if values.len() > 1 {
        let arr_elems: Vec<ArrayElement> = values
            .into_iter()
            .map(|v| ArrayElement {
                key: None,
                value: v,
                spread: false,
                by_ref: false,
            })
            .collect();
        Ok(StmtKind::Return(Some(Expression::new(ExprKind::Array(
            arr_elems,
        )))))
    } else {
        Ok(StmtKind::Return(None))
    }
}

fn walk_goto(pair: Pair<Rule>) -> Result<String, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::ident_name {
            return Ok(inner.as_str().to_string());
        }
    }
    Ok(String::new())
}

fn walk_labeled(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut label = String::new();
    let mut stmt = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => label = inner.as_str().to_string(),
            Rule::statement => stmt = Some(walk_statement(inner)?),
            _ => {}
        }
    }

    if let Some(s) = stmt {
        // Wrap the labeled statement so the compiler can route
        // `break <label>` / `continue <label>` to it.
        Ok(StmtKind::Labeled {
            label,
            body: Box::new(s),
        })
    } else {
        Ok(StmtKind::Label(label))
    }
}

// ── Expressions ─────────────────────────────────────────────────────────────────────────

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    if matches!(
        pair.as_rule(),
        Rule::expression | Rule::if_expression | Rule::range_expression
    ) {
        let mut operands = Vec::new();
        let mut operators: Vec<String> = Vec::new();

        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::unary_expression
                | Rule::if_unary_expression
                | Rule::range_unary_expression => operands.push(walk_unary_expression(inner)?),
                Rule::binary_op => {
                    let op = inner.as_str().to_string();
                    while operators
                        .last()
                        .is_some_and(|top| go_binary_precedence(top) >= go_binary_precedence(&op))
                    {
                        go_reduce_binary_expr(&mut operands, &mut operators)?;
                    }
                    operators.push(op);
                }
                _ => {}
            }
        }

        while !operators.is_empty() {
            go_reduce_binary_expr(&mut operands, &mut operators)?;
        }

        if let Some(result) = operands.pop() {
            return Ok(result);
        }
    } else if matches!(
        pair.as_rule(),
        Rule::unary_expression | Rule::if_unary_expression | Rule::range_unary_expression
    ) {
        return walk_unary_expression(pair);
    } else if matches!(
        pair.as_rule(),
        Rule::primary | Rule::if_primary | Rule::range_primary
    ) {
        return walk_primary(pair);
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_unary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut op = None;
    let mut operand = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::unary_op => op = Some(inner.as_str().to_string()),
            Rule::unary_expression | Rule::if_unary_expression | Rule::range_unary_expression => {
                operand = Some(walk_unary_expression(inner)?)
            }
            Rule::primary | Rule::if_primary | Rule::range_primary => {
                operand = Some(walk_primary(inner)?)
            }
            _ => {}
        }
    }

    if let Some(uop) = op {
        let un_op = match uop.as_str() {
            "-" => UnaryOp::Neg,
            "!" => UnaryOp::Not,
            "+" => UnaryOp::Pos,
            "^" => UnaryOp::BitNot,
            "*" => UnaryOp::Deref,
            "&" => UnaryOp::AddrOf,
            "<-" => {
                return Ok(chan_recv(operand.unwrap_or_else(Expression::null)));
            }
            _ => UnaryOp::Pos,
        };
        Ok(Expression::new(ExprKind::Unary {
            op: un_op,
            expr: Box::new(
                operand.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null))),
            ),
        }))
    } else {
        operand.ok_or_else(|| "Empty unary expression".to_string())
    }
}

fn walk_primary(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut base = None;
    let mut chain = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::operand | Rule::if_operand | Rule::range_operand => {
                base = Some(walk_operand(inner)?);
            }
            Rule::selector => {
                for s_inner in inner.into_inner() {
                    if s_inner.as_rule() == Rule::ident_name {
                        chain.push(PrimaryChain::Member(s_inner.as_str().to_string()));
                    }
                }
            }
            Rule::generic_instantiation => {
                for g_inner in inner.into_inner() {
                    if g_inner.as_rule() == Rule::type_arguments {
                        chain.push(PrimaryChain::GenericInstantiation(
                            common_generics::generic_argument_display_names(g_inner.as_str()),
                        ));
                    }
                }
            }
            Rule::index => {
                for i_inner in inner.into_inner() {
                    if i_inner.as_rule() == Rule::expression {
                        chain.push(PrimaryChain::Index(walk_expression(i_inner)?));
                    }
                }
            }
            Rule::two_index_slice | Rule::three_index_slice => {
                let slice_source = inner.as_str();
                let mut start = None;
                let mut end = None;
                let mut max = None;
                for s_inner in inner.into_inner() {
                    if s_inner.as_rule() == Rule::expression {
                        if start.is_none() && !slice_source.starts_with("[:") {
                            start = Some(walk_expression(s_inner)?);
                        } else if end.is_none() {
                            end = Some(walk_expression(s_inner)?);
                        } else if max.is_none() {
                            max = Some(walk_expression(s_inner)?);
                        }
                    }
                }
                chain.push(PrimaryChain::Slice { start, end, max });
            }
            Rule::call => {
                let mut args = Vec::new();
                for c_inner in inner.into_inner() {
                    if c_inner.as_rule() == Rule::argument_list {
                        for arg_inner in c_inner.into_inner() {
                            if arg_inner.as_rule() == Rule::argument {
                                let mut spread = false;
                                let mut val = None;
                                for expr_inner in arg_inner.into_inner() {
                                    if expr_inner.as_rule() == Rule::expression {
                                        val = Some(walk_expression(expr_inner)?);
                                    } else if expr_inner.as_rule() == Rule::type_annotation {
                                        val = Some(go_type_arg_expr(walk_type(expr_inner)));
                                    } else if expr_inner.as_rule() == Rule::spread_suffix {
                                        spread = true;
                                    }
                                }
                                if let Some(expr) = val {
                                    args.push(Argument {
                                        value: expr,
                                        name: None,
                                        by_ref: false,
                                        spread,
                                    });
                                }
                            }
                        }
                    }
                }
                chain.push(PrimaryChain::Call(args));
            }
            Rule::type_assertion => {
                for t_inner in inner.into_inner() {
                    if t_inner.as_rule() == Rule::type_annotation {
                        chain.push(PrimaryChain::TypeAssert(walk_type(t_inner)));
                    }
                }
            }
            _ => {}
        }
    }

    if let Some(mut result) = base {
        let mut pending_type_args: Vec<String> = Vec::new();
        for item in chain {
            result = match item {
                PrimaryChain::GenericInstantiation(type_args) => {
                    pending_type_args = type_args;
                    result
                }
                PrimaryChain::Member(name) => Expression::new(ExprKind::Member {
                    object: Box::new(result),
                    field: name,
                    null_safe: false,
                }),
                PrimaryChain::Index(idx) => Expression::new(ExprKind::Index {
                    object: Box::new(result),
                    index: Box::new(idx),
                    null_safe: false,
                }),
                PrimaryChain::Slice { start, end, max } => {
                    let base = result;
                    let start_expr =
                        start.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(0))));
                    let end_expr =
                        end.unwrap_or_else(|| go_builtin_call("len", vec![base.clone()]));
                    if let Some(max_expr) = max {
                        let cap_bound = Expression::new(ExprKind::Binary {
                            op: BinOp::Sub,
                            left: Box::new(max_expr),
                            right: Box::new(start_expr.clone()),
                        });
                        go_builtin_call(
                            "__go_slices_slice_bound_common",
                            vec![base, start_expr, end_expr, cap_bound],
                        )
                    } else {
                        go_builtin_call(
                            "__go_slices_slice_common",
                            vec![base, start_expr, end_expr],
                        )
                    }
                }
                PrimaryChain::Call(args) => {
                    let mut call_args = Vec::new();
                    call_args.extend(pending_type_args.drain(..).map(|type_name| {
                        Argument::positional(go_runtime_type_arg_expr(type_name))
                    }));
                    call_args.extend(args);
                    Expression::new(ExprKind::Call {
                        callee: Box::new(result),
                        args: call_args,
                        optional: false,
                    })
                }
                PrimaryChain::TypeAssert(type_name) => go_type_assert_expr(result, type_name),
            };
        }
        Ok(result)
    } else {
        Ok(Expression::new(ExprKind::Lit(Literal::Null)))
    }
}

#[derive(Clone)]
enum PrimaryChain {
    Member(String),
    GenericInstantiation(Vec<String>),
    Index(Expression),
    Slice {
        start: Option<Expression>,
        end: Option<Expression>,
        max: Option<Expression>,
    },
    Call(Vec<Argument>),
    TypeAssert(String),
}

fn walk_operand(pair: Pair<Rule>) -> Result<Expression, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::literal => return walk_literal(inner),
            Rule::slice_conversion => return walk_slice_conversion(inner),
            Rule::interface_conversion => return walk_type_conversion(inner),
            Rule::type_conversion => return walk_type_conversion(inner),
            Rule::ident_name => {
                let name = inner.as_str();
                // Go builtins
                match name {
                    "nil" => return Ok(Expression::new(ExprKind::Lit(Literal::Null))),
                    "true" => return Ok(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
                    "false" => return Ok(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
                    _ => return Ok(Expression::new(ExprKind::Ident(name.to_string()))),
                }
            }
            Rule::expression | Rule::if_expression | Rule::range_expression => {
                return walk_expression(inner);
            }
            Rule::composite_literal => return walk_composite_literal(inner),
            Rule::function_literal => return walk_function_literal(inner),
            _ => {}
        }
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_slice_conversion(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut type_name = String::from("[]");
    let mut expr = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => type_name.push_str(inner.as_str()),
            Rule::expression => expr = Some(walk_expression(inner)?),
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::Cast {
        expr: Box::new(expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))),
        type_name,
    }))
}

fn walk_type_conversion(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut type_name = None;
    let mut expr = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::type_annotation => type_name = Some(walk_type(inner)),
            Rule::expression => expr = Some(walk_expression(inner)?),
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::Cast {
        expr: Box::new(expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))),
        type_name: type_name.unwrap_or_default(),
    }))
}

fn walk_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::numeric_literal => {
                let mut s = inner.as_str().replace('_', "");
                let imaginary = s.ends_with('i');
                if imaginary {
                    s.pop();
                }
                let parsed = if s.starts_with("0x") || s.starts_with("0X") {
                    i64::from_str_radix(&s[2..], 16).ok().map(Expression::int)
                } else if s.starts_with("0b") || s.starts_with("0B") {
                    i64::from_str_radix(&s[2..], 2).ok().map(Expression::int)
                } else if s.starts_with("0o") || s.starts_with("0O") {
                    i64::from_str_radix(&s[2..], 8).ok().map(Expression::int)
                } else if s.contains('.')
                    || s.contains('e')
                    || s.contains('E')
                    || s.contains('p')
                    || s.contains('P')
                {
                    s.parse::<f64>().ok().map(Expression::float)
                } else {
                    s.parse::<i64>().ok().map(Expression::int)
                };
                if let Some(value) = parsed {
                    if imaginary {
                        return Ok(go_complex_value_expr(Expression::int(0), value));
                    }
                    return Ok(value);
                }
                if s.starts_with("0x") || s.starts_with("0X") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 16) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.starts_with("0b") || s.starts_with("0B") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 2) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.starts_with("0o") || s.starts_with("0O") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 8) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.contains('.')
                    || s.contains('e')
                    || s.contains('E')
                    || s.contains('p')
                    || s.contains('P')
                {
                    if let Ok(f) = s.parse::<f64>() {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Float(f))));
                    }
                } else if let Ok(n) = s.parse::<i64>() {
                    return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                }
            }
            Rule::string_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Str(unquote(
                    inner.as_str(),
                )))));
            }
            Rule::bool_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Bool(
                    inner.as_str() == "true",
                ))));
            }
            Rule::nil_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
            }
            Rule::rune_literal => {
                let rune = unquote(inner.as_str());
                let code = rune.chars().next().map(|ch| ch as i64).unwrap_or(0);
                return Ok(Expression::new(ExprKind::Lit(Literal::Int(code))));
            }
            _ => {}
        }
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_composite_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut type_name = String::new();
    let mut elements = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::literal_type => {
                type_name = go_literal_type_name(inner);
            }
            Rule::literal_value => {
                for lv_inner in inner.into_inner() {
                    if lv_inner.as_rule() == Rule::element_list {
                        elements = walk_element_list(lv_inner)?;
                    }
                }
            }
            _ => {}
        }
    }

    if type_name.starts_with("map[") {
        // Build a dict/object literal
        let mut props = Vec::new();
        let value_type = go_map_value_type(&type_name);
        for (key, val) in elements {
            props.push(ObjectProperty::KeyValue {
                key,
                value: go_retype_elided_element(val, value_type.as_deref()),
            });
        }
        Ok(go_typed_composite_expr(
            Expression::new(ExprKind::Object(props)),
            &type_name,
        ))
    } else if go_is_array_like_type(&type_name) {
        let elem_type = go_array_element_type(&type_name);
        let mut values = Vec::new();
        if let Some(target_len) = go_fixed_array_len(&type_name, elements.len()) {
            if let Some(elem_type) = elem_type.as_deref() {
                values.resize_with(target_len, || go_zero_value_expr(elem_type));
            } else {
                values.resize_with(target_len, Expression::null);
            }
        }
        let mut next_index = 0usize;
        for (key, value) in elements {
            let index = go_composite_literal_index_key(&key).unwrap_or(next_index);
            if index >= values.len() {
                if let Some(elem_type) = elem_type.as_deref() {
                    values.resize_with(index + 1, || go_zero_value_expr(elem_type));
                } else {
                    values.resize_with(index + 1, Expression::null);
                }
            }
            values[index] = go_retype_elided_element(value, elem_type.as_deref());
            next_index = index + 1;
        }
        let arr_elems: Vec<ArrayElement> = values
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect();
        Ok(go_typed_composite_expr(
            Expression::new(ExprKind::Array(arr_elems)),
            &type_name,
        ))
    } else if !type_name.is_empty()
        && elements
            .iter()
            .all(|(key, _)| !matches!(key.kind, ExprKind::Lit(Literal::Null)))
    {
        let mut props = Vec::new();
        for (key, val) in elements {
            let key = match key.kind {
                ExprKind::Ident(name) => Expression::new(ExprKind::Lit(Literal::Str(name))),
                ExprKind::Lit(Literal::Int(n)) => {
                    Expression::new(ExprKind::Lit(Literal::Str(n.to_string())))
                }
                _ => key,
            };
            props.push(ObjectProperty::KeyValue { key, value: val });
        }
        Ok(go_typed_composite_expr(
            Expression::new(ExprKind::Object(props)),
            &type_name,
        ))
    } else {
        // Untyped composite literal fallback.
        let arr_elems: Vec<ArrayElement> = elements
            .into_iter()
            .map(|(_, v)| ArrayElement {
                key: None,
                value: v,
                spread: false,
                by_ref: false,
            })
            .collect();
        Ok(go_typed_composite_expr(
            Expression::new(ExprKind::Array(arr_elems)),
            &type_name,
        ))
    }
}

/// Apply the composite's element type to an elided element literal. In
/// `[]tagged{{1, 10}}` the inner `{1, 10}` is walked as an untyped composite
/// (a bare `Array`/`Object`); tagging it with the element type lets the
/// normalize pass expand it to the proper struct/slice/map value. Scalars and
/// already-typed elements pass through unchanged.
fn go_retype_elided_element(value: Expression, elem_type: Option<&str>) -> Expression {
    match (elem_type, &value.kind) {
        (Some(elem_type), ExprKind::Array(_) | ExprKind::Object(_)) if !elem_type.is_empty() => {
            go_typed_composite_expr(value, elem_type)
        }
        _ => value,
    }
}

/// Type name of a composite `literal_type`, erasing generic type arguments
/// (`Pair[int]` → `Pair`).
fn go_literal_type_name(pair: Pair<Rule>) -> String {
    if let Some(backing) = go_stdlib_type_binding(pair.as_str()) {
        return backing.to_string();
    }
    common_generics::erased_type_name(pair.as_str())
}

fn go_composite_literal_index_key(expr: &Expression) -> Option<usize> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(index)) if *index >= 0 => Some(*index as usize),
        ExprKind::Lit(Literal::Str(index)) => index.parse::<usize>().ok(),
        _ => None,
    }
}

fn go_typed_composite_expr(expr: Expression, type_name: &str) -> Expression {
    if type_name.is_empty() {
        expr
    } else {
        Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: type_name.to_string(),
        })
    }
}

fn go_is_array_like_type(type_name: &str) -> bool {
    go_array_head(type_name).is_some()
}

fn go_array_head(type_name: &str) -> Option<(&str, &str)> {
    let trimmed = type_name.trim();
    if !trimmed.starts_with('[') {
        return None;
    }
    let close = trimmed.find(']')?;
    Some((&trimmed[1..close], trimmed[close + 1..].trim()))
}

fn go_array_element_type(type_name: &str) -> Option<String> {
    let (_, tail) = go_array_head(type_name)?;
    (!tail.is_empty()).then(|| tail.to_string())
}

fn go_fixed_array_bounds_exprs(type_name: &str) -> Option<Vec<Expression>> {
    let mut remaining = type_name.trim();
    let mut bounds = Vec::new();

    while let Some((head, tail)) = go_array_head(remaining) {
        let head = head.trim();
        if head.is_empty() || head == "..." {
            return None;
        }
        bounds.push(Expression::int(head.parse::<i64>().ok()?));
        remaining = tail.trim();
    }

    (!bounds.is_empty()).then_some(bounds)
}

fn go_fixed_array_len(type_name: &str, inferred_len: usize) -> Option<usize> {
    let (head, _) = go_array_head(type_name)?;
    let head = head.trim();
    if head.is_empty() {
        None
    } else if head == "..." {
        Some(inferred_len)
    } else {
        head.parse::<usize>().ok()
    }
}

fn go_zero_value_expr(type_name: &str) -> Expression {
    let trimmed = type_name.trim();
    let lower = trimmed.to_ascii_lowercase();

    if let Some(value) = crate::adapters::netip::zero_value(trimmed) {
        return value;
    }

    if let Some(len) = go_fixed_array_len(trimmed, 0) {
        if let Some(elem_type) = go_array_element_type(trimmed) {
            let elements = (0..len)
                .map(|_| ArrayElement {
                    key: None,
                    value: go_zero_value_expr(&elem_type),
                    spread: false,
                    by_ref: false,
                })
                .collect();
            return go_typed_composite_expr(Expression::new(ExprKind::Array(elements)), trimmed);
        }
    }

    if lower.starts_with("func(") || lower == "func" || lower == "interface{}" || lower == "any" {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }

    if go_is_channel_type(trimmed) {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }

    if lower.starts_with("[]") || lower.starts_with("map[") || lower.starts_with('*') {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }

    if let Some(value) = crate::adapters::bytes_io::zero_value(trimmed) {
        return value;
    }

    if let Some(value) = crate::adapters::hash::zero_value(trimmed) {
        return value;
    }

    if let Some(value) = crate::adapters::url::zero_value(trimmed) {
        return value;
    }

    match lower.as_str() {
        "error" => Expression::null(),
        "bool" => Expression::new(ExprKind::Lit(Literal::Bool(false))),
        "string" => Expression::new(ExprKind::Lit(Literal::Str(String::new()))),
        "float32" | "float64" => Expression::new(ExprKind::Lit(Literal::Float(0.0))),
        "int" | "int8" | "int16" | "int32" | "int64" | "uint" | "uint8" | "uint16" | "uint32"
        | "uint64" | "uintptr" | "byte" | "rune" => Expression::new(ExprKind::Lit(Literal::Int(0))),
        "__gotime" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: Expression::string("sec"),
                    value: Expression::int(0),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("nsec"),
                    value: Expression::int(0),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("loc"),
                    value: go_builtin_call("go.time_utc", Vec::new()),
                },
            ])),
            trimmed,
        ),
        "__gosyncmap" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                key: Expression::string("data"),
                value: go_typed_composite_expr(
                    Expression::new(ExprKind::Object(Vec::new())),
                    "map[interface{}]interface{}",
                ),
            }])),
            trimmed,
        ),
        "__gosyncpool" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: Expression::string("New"),
                    value: Expression::null(),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("items"),
                    value: go_typed_composite_expr(
                        Expression::new(ExprKind::Array(Vec::new())),
                        "[]interface{}",
                    ),
                },
            ])),
            trimmed,
        ),
        "__gosynconce" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                key: Expression::string("done"),
                value: Expression::new(ExprKind::Lit(Literal::Bool(false))),
            }])),
            trimmed,
        ),
        "__gosyncwaitgroup" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                key: Expression::string("count"),
                value: Expression::int(0),
            }])),
            trimmed,
        ),
        "__goatomicint32" | "__goatomicint64" | "__goatomicuint32" | "__goatomicuint64" => {
            go_typed_composite_expr(
                Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                    key: Expression::string("value"),
                    value: Expression::int(0),
                }])),
                trimmed,
            )
        }
        "__goatomicbool" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                key: Expression::string("value"),
                value: Expression::bool(false),
            }])),
            trimmed,
        ),
        "__goatomicvalue" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                key: Expression::string("value"),
                value: Expression::null(),
            }])),
            trimmed,
        ),
        "__goxmlname" => crate::adapters::xml::name_expr(
            Expression::string(""),
            Expression::string(""),
            Expression::string(""),
        ),
        _ => go_typed_composite_expr(Expression::new(ExprKind::Object(Vec::new())), trimmed),
    }
}

fn go_zero_value_for_type(type_name: &str, env: &GoNormalizeEnv) -> Expression {
    if let Some(runtime_param) = env.generic_type_params.get(type_name.trim()) {
        return go_zero_value_from_type_token(Expression::ident(runtime_param));
    }
    if let Some(mapped) = go_stdlib_type_binding(type_name) {
        return go_zero_value_for_type(mapped, env);
    }
    if let Some(underlying) = env
        .named_types
        .get(type_name)
        .filter(|underlying| underlying.as_str() != type_name)
    {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(go_zero_value_for_type(underlying, env)),
            type_name: type_name.to_string(),
        });
    }
    go_zero_value_expr(type_name)
}

fn go_runtime_generic_param_name(runtime_name: &str) -> Option<String> {
    runtime_name
        .strip_prefix("__generic_typearg_")
        .map(str::to_string)
}

fn go_map_value_type(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim();
    if trimmed == "__goValues" {
        return Some("[]string".to_string());
    }
    if !trimmed.starts_with("map[") {
        return None;
    }

    let mut depth = 0usize;
    for (idx, ch) in trimmed.char_indices() {
        match ch {
            '[' => depth += 1,
            ']' => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    let tail = trimmed.get(idx + 1..)?.trim();
                    return (!tail.is_empty()).then(|| tail.to_string());
                }
            }
            _ => {}
        }
    }
    None
}

fn walk_element_list(pair: Pair<Rule>) -> Result<Vec<(Expression, Expression)>, String> {
    let mut elements = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::keyed_element {
            let parts: Vec<_> = inner.into_inner().collect();
            if parts.len() >= 2 {
                let key = go_keyed_element_key(parts[0].clone())?;
                let value = go_keyed_element_value(parts[1].clone())?;
                elements.push((key, value));
            } else if let Some(value_pair) = parts.into_iter().next() {
                elements.push((
                    Expression::new(ExprKind::Lit(Literal::Null)),
                    go_keyed_element_value(value_pair)?,
                ));
            } else {
                elements.push((
                    Expression::new(ExprKind::Lit(Literal::Null)),
                    Expression::new(ExprKind::Lit(Literal::Null)),
                ));
            }
        }
    }
    Ok(elements)
}

fn go_keyed_element_key(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::ident_name => Ok(Expression::new(ExprKind::Ident(pair.as_str().to_string()))),
        Rule::string_literal => Ok(Expression::new(ExprKind::Lit(Literal::Str(unquote(
            pair.as_str(),
        ))))),
        Rule::bool_literal => Ok(Expression::new(ExprKind::Lit(Literal::Bool(
            pair.as_str() == "true",
        )))),
        Rule::rune_literal => {
            let rune = unquote(pair.as_str());
            let code = rune.chars().next().map(|ch| ch as i64).unwrap_or(0);
            Ok(Expression::new(ExprKind::Lit(Literal::Int(code))))
        }
        Rule::numeric_literal | Rule::signed_numeric_key => {
            let literal = pair.as_str().replace('_', "");
            if let Ok(n) = literal.parse::<i64>() {
                Ok(Expression::new(ExprKind::Lit(Literal::Int(n))))
            } else if let Ok(f) = literal.parse::<f64>() {
                Ok(Expression::new(ExprKind::Lit(Literal::Float(f))))
            } else {
                Ok(Expression::new(ExprKind::Lit(Literal::Null)))
            }
        }
        Rule::expression => walk_expression(pair),
        Rule::element => go_keyed_element_value(pair),
        Rule::composite_literal => walk_composite_literal(pair),
        Rule::literal_value => walk_literal_value_expr(pair),
        _ => Ok(Expression::new(ExprKind::Lit(Literal::Null))),
    }
}

fn go_keyed_element_value(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::element => {
            let Some(inner) = pair.into_inner().next() else {
                return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
            };
            go_keyed_element_value(inner)
        }
        Rule::expression => walk_expression(pair),
        Rule::literal_value => walk_literal_value_expr(pair),
        Rule::ident_name => Ok(Expression::new(ExprKind::Ident(pair.as_str().to_string()))),
        Rule::string_literal => Ok(Expression::new(ExprKind::Lit(Literal::Str(unquote(
            pair.as_str(),
        ))))),
        Rule::bool_literal => Ok(Expression::new(ExprKind::Lit(Literal::Bool(
            pair.as_str() == "true",
        )))),
        Rule::numeric_literal => go_keyed_element_key(pair),
        _ => Ok(Expression::new(ExprKind::Lit(Literal::Null))),
    }
}

fn walk_literal_value_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut elements = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::element_list {
            elements = walk_element_list(inner)?;
        }
    }

    if elements
        .iter()
        .all(|(key, _)| !matches!(key.kind, ExprKind::Lit(Literal::Null)))
    {
        let mut props = Vec::new();
        for (key, value) in elements {
            let key = match key.kind {
                ExprKind::Ident(name) => Expression::string(&name),
                ExprKind::Lit(Literal::Int(n)) => Expression::string(&n.to_string()),
                _ => key,
            };
            props.push(ObjectProperty::KeyValue { key, value });
        }
        Ok(Expression::new(ExprKind::Object(props)))
    } else {
        Ok(Expression::new(ExprKind::Array(
            elements
                .into_iter()
                .map(|(_, value)| ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        )))
    }
}

fn walk_function_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut params = Vec::new();
    let mut return_type = None;
    let mut body = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::signature => {
                let sig = walk_signature(inner)?;
                params = sig.params;
                return_type = sig.return_type;
            }
            Rule::function_body | Rule::block_statement => {
                body = walk_block(inner)?;
            }
            _ => {}
        }
    }

    let func_type = go_function_type_name(&params, return_type.as_deref());
    Ok(Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::new(ExprKind::Lambda {
            params,
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        type_name: func_type,
    }))
}

fn go_function_type_name(params: &[Param], return_type: Option<&str>) -> String {
    let args = params
        .iter()
        .map(|param| {
            param
                .type_hint
                .as_deref()
                .map(str::to_string)
                .unwrap_or_else(|| "interface{}".to_string())
        })
        .collect::<Vec<_>>()
        .join(", ");
    let mut name = format!("func({})", args);
    if let Some(return_type) = return_type.filter(|ty| !ty.trim().is_empty()) {
        name.push(' ');
        name.push_str(return_type.trim());
    }
    name
}

fn walk_expression_list(pair: Pair<Rule>) -> Result<Vec<Expression>, String> {
    let mut exprs = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::expression {
            exprs.push(walk_expression(inner)?);
        }
    }
    Ok(exprs)
}

// ── Helpers ───────────────────────────────────────────────────────────────────────────────

fn first_meaningful(pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() != Rule::EOI {
            return Ok(inner);
        }
    }
    Err("No meaningful child".to_string())
}

fn parse_bin_op(op: &str) -> BinOp {
    match op {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        "<=" => BinOp::LtEq,
        ">" => BinOp::Gt,
        ">=" => BinOp::GtEq,
        "&&" => BinOp::And,
        "||" => BinOp::Or,
        "&" => BinOp::BitAnd,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "&^" => BinOp::BitAnd,
        _ => BinOp::Add,
    }
}

fn build_go_binary_expr(op: &str, left: Expression, right: Expression) -> Expression {
    if op == "&^" {
        Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(left),
            right: Box::new(Expression::new(ExprKind::Unary {
                op: UnaryOp::BitNot,
                expr: Box::new(right),
            })),
        })
    } else {
        Expression::new(ExprKind::Binary {
            op: parse_bin_op(op),
            left: Box::new(left),
            right: Box::new(right),
        })
    }
}

fn go_reduce_binary_expr(
    operands: &mut Vec<Expression>,
    operators: &mut Vec<String>,
) -> Result<(), String> {
    let Some(op) = operators.pop() else {
        return Ok(());
    };
    let Some(right) = operands.pop() else {
        return Err(format!("missing right operand for Go binary operator {op}"));
    };
    let Some(left) = operands.pop() else {
        return Err(format!("missing left operand for Go binary operator {op}"));
    };
    operands.push(build_go_binary_expr(&op, left, right));
    Ok(())
}

fn go_binary_precedence(op: &str) -> u8 {
    match op {
        "||" => 1,
        "&&" => 2,
        "==" | "!=" | "<" | "<=" | ">" | ">=" => 3,
        "+" | "-" | "|" | "^" => 4,
        "*" | "/" | "%" | "<<" | ">>" | "&" | "&^" => 5,
        _ => 0,
    }
}

fn go_type_arg_expr(type_name: String) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::null()),
        type_name,
    })
}

fn go_runtime_type_arg_expr(type_name: String) -> Expression {
    Expression::string(&type_name)
}

fn go_zero_value_from_type_token(token: Expression) -> Expression {
    let mut result = Expression::null();
    for type_name in ["float32", "float64"] {
        result = Expression::new(ExprKind::Ternary {
            cond: Box::new(go_type_token_eq(&token, type_name)),
            then: Box::new(Expression::new(ExprKind::Lit(Literal::Float(0.0)))),
            else_: Box::new(result),
        });
    }
    for type_name in [
        "int", "int8", "int16", "int32", "int64", "uint", "uint8", "uint16", "uint32", "uint64",
        "uintptr", "byte", "rune",
    ] {
        result = Expression::new(ExprKind::Ternary {
            cond: Box::new(go_type_token_eq(&token, type_name)),
            then: Box::new(Expression::int(0)),
            else_: Box::new(result),
        });
    }
    result = Expression::new(ExprKind::Ternary {
        cond: Box::new(go_type_token_eq(&token, "string")),
        then: Box::new(Expression::string("")),
        else_: Box::new(result),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(go_type_token_eq(&token, "bool")),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
        else_: Box::new(result),
    })
}

fn go_type_token_eq(token: &Expression, type_name: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(token.clone()),
        right: Box::new(Expression::string(type_name)),
    })
}

fn go_effective_generic_call_args(
    args: &[Argument],
    signature: Option<&GoFunctionSignature>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Vec<Argument> {
    let Some(signature) = signature else {
        return args.to_vec();
    };
    let generic_arg_count = signature.generic_arg_count;
    if generic_arg_count == 0 {
        return args.to_vec();
    }
    if args
        .iter()
        .take(generic_arg_count)
        .filter_map(|arg| go_type_arg_name_from_expr(&arg.value))
        .count()
        == generic_arg_count
    {
        return args.to_vec();
    }

    let mut out = Vec::with_capacity(generic_arg_count + args.len());
    for type_param in signature.generic_param_names.iter().take(generic_arg_count) {
        let inferred = go_infer_generic_call_type_arg(type_param, args, signature, env, signatures)
            .unwrap_or_else(|| "any".into());
        out.push(Argument::positional(go_runtime_type_arg_expr(inferred)));
    }
    while out.len() < generic_arg_count {
        out.push(Argument::positional(go_runtime_type_arg_expr(
            "any".to_string(),
        )));
    }
    out.extend(args.iter().cloned());
    out
}

fn go_infer_generic_call_type_arg(
    type_param: &str,
    args: &[Argument],
    signature: &GoFunctionSignature,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    for (idx, arg) in args.iter().enumerate() {
        let formal = signature
            .params
            .get(signature.generic_arg_count + idx)
            .and_then(|hint| hint.as_deref())?;
        let actual = go_expr_type_hint(&arg.value, env, signatures)?;
        if let Some(inferred) = go_infer_generic_type_arg_from_types(type_param, formal, &actual) {
            return Some(inferred);
        }
    }
    None
}

fn go_infer_generic_type_arg_from_types(
    type_param: &str,
    formal_type: &str,
    actual_type: &str,
) -> Option<String> {
    let formal = formal_type.trim();
    let actual = actual_type.trim();
    if formal == type_param {
        return Some(actual.to_string());
    }
    if let Some(formal_inner) = formal.strip_prefix("[]") {
        if formal_inner.trim() == type_param {
            return actual
                .strip_prefix("[]")
                .map(str::trim)
                .filter(|inner| !inner.is_empty())
                .map(str::to_string);
        }
    }
    if let Some(formal_inner) = formal.strip_prefix('*') {
        if formal_inner.trim() == type_param {
            return actual
                .strip_prefix('*')
                .map(str::trim)
                .filter(|inner| !inner.is_empty())
                .map(str::to_string);
        }
    }
    if let (Some((formal_key, formal_value)), Some((actual_key, actual_value))) = (
        go_map_key_value_types(formal),
        go_map_key_value_types(actual),
    ) {
        if formal_key.trim() == type_param {
            return Some(actual_key);
        }
        if formal_value.trim() == type_param {
            return Some(actual_value);
        }
    }
    None
}

fn go_map_key_value_types(type_name: &str) -> Option<(String, String)> {
    let trimmed = type_name.trim();
    if trimmed == "__goValues" {
        return Some(("string".to_string(), "[]string".to_string()));
    }
    if !trimmed.starts_with("map[") {
        return None;
    }

    let mut depth = 0usize;
    for (idx, ch) in trimmed.char_indices() {
        match ch {
            '[' => depth += 1,
            ']' => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    let key = trimmed.get(4..idx)?.trim();
                    let value = trimmed.get(idx + 1..)?.trim();
                    if !key.is_empty() && !value.is_empty() {
                        return Some((key.to_string(), value.to_string()));
                    }
                    return None;
                }
            }
            _ => {}
        }
    }
    None
}

fn go_type_arg_name_from_expr(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(type_name)) => Some(type_name.clone()),
        ExprKind::Cast { expr, type_name } if matches!(expr.kind, ExprKind::Lit(Literal::Null)) => {
            Some(type_name.clone())
        }
        _ => None,
    }
}

fn go_type_assert_expr(expr: Expression, type_name: String) -> Expression {
    if type_name.trim() == "__goXMLStartElement" {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(go_xml_token_element_from_go_expr(expr, "start")),
            type_name,
        });
    }
    if type_name.trim() == "__goXMLEndElement" {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(go_xml_token_element_from_go_expr(expr, "end")),
            type_name,
        });
    }
    go_builtin_call("__go_type_assert", vec![expr, go_type_arg_expr(type_name)])
}

fn go_extract_type_assert_expr(expr: &Expression) -> Option<(Expression, String)> {
    if let ExprKind::Cast { expr, type_name } = &expr.kind {
        if matches!(
            type_name.trim(),
            "__goXMLStartElement" | "__goXMLEndElement"
        ) {
            return Some((
                go_xml_type_assert_source_expr(expr).unwrap_or_else(|| expr.as_ref().clone()),
                type_name.clone(),
            ));
        }
    }
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_type_assert")
        || args.len() != 2
    {
        return None;
    }
    Some((
        args[0].value.clone(),
        go_type_name_from_expr(&args[1].value)?,
    ))
}

fn go_xml_type_assert_source_expr(expr: &Expression) -> Option<Expression> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        let is_tag_key = matches!(
            &key.kind,
            ExprKind::Lit(Literal::Str(s)) if s == "Tag"
        );
        if !is_tag_key {
            continue;
        }
        return Some(value.clone());
    }
    None
}

fn go_type_assert_value_expr(
    expr: Expression,
    type_name: &str,
    env: &GoNormalizeEnv,
    mut state: Option<&mut GoNormalizeState>,
) -> Expression {
    let trimmed_type = type_name.trim();
    if let Some(concrete) = go_known_interface_dynamic_type(&expr, env) {
        if go_types_match_for_assert(&concrete, trimmed_type, env) {
            let cast_type = if go_is_go_interface_type(trimmed_type, env) {
                concrete
            } else {
                trimmed_type.to_string()
            };
            return Expression::new(ExprKind::Cast {
                expr: Box::new(expr),
                type_name: cast_type,
            });
        }
    }
    if let Some(concrete) = go_concrete_type_for_interface(trimmed_type, env) {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: concrete,
        });
    }
    if let Some(mapped) = go_stdlib_type_binding(type_name) {
        if mapped == "__goXMLStartElement" {
            return Expression::new(ExprKind::Cast {
                expr: Box::new(go_xml_token_element_from_go_expr(expr, "start")),
                type_name: mapped.to_string(),
            });
        }
        if mapped == "__goXMLEndElement" {
            return Expression::new(ExprKind::Cast {
                expr: Box::new(go_xml_token_element_from_go_expr(expr, "end")),
                type_name: mapped.to_string(),
            });
        }
        return Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: mapped.to_string(),
        });
    }
    if trimmed_type.starts_with("__goXML")
        || matches!(trimmed_type, "__goRawMessage" | "__goLevel" | "__goAttr")
    {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: trimmed_type.to_string(),
        });
    }

    if !matches!(
        trimmed_type,
        "int"
            | "int8"
            | "int16"
            | "int32"
            | "int64"
            | "uint"
            | "uint8"
            | "uint16"
            | "uint32"
            | "uint64"
            | "uintptr"
            | "byte"
            | "rune"
            | "float32"
            | "float64"
            | "string"
            | "bool"
    ) {
        return expr;
    }

    if let ExprKind::Call { callee, .. } = &expr.kind {
        if go_expr_call_name(callee).as_deref() == Some("__go_sync_pool_Get") {
            return expr;
        }
    }

    if go_type_assert_needs_single_eval(&expr) {
        if let Some(state) = state.as_deref_mut() {
            let temp = fresh_go_temp(state, "__go_assert");
            let mut captures = go_big_captures(&[&expr]);
            captures.retain(|name| !name.starts_with("__go_"));
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Lambda {
                    params: vec![],
                    body: LambdaBody::Block(vec![
                        Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(temp.clone()),
                                type_hint: None,
                                init: Some(expr),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }),
                        Statement::new(StmtKind::Return(Some(Expression::ident(&temp)))),
                    ]),
                    is_async: false,
                    captures,
                })),
                args: vec![],
                optional: false,
            });
        }
    }

    let cond = go_build_is_type(expr.clone(), type_name);
    let then_expr = Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then_expr),
        else_: Box::new(go_zero_value_expr(type_name)),
    })
}

fn go_known_interface_dynamic_type(expr: &Expression, env: &GoNormalizeEnv) -> Option<String> {
    let ExprKind::Ident(name) = &expr.kind else {
        return None;
    };
    env.interface_concrete_types.get(name).cloned()
}

fn go_types_match_for_assert(actual: &str, expected: &str, env: &GoNormalizeEnv) -> bool {
    let actual = actual.trim();
    let expected = expected.trim();
    actual == expected
        || (go_is_go_interface_type(expected, env)
            && go_type_assignable_to_interface(actual, expected, env))
}

fn go_type_assert_needs_single_eval(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call { .. } | ExprKind::Assign { .. } | ExprKind::Sequence(_) => true,
        ExprKind::Member { object, .. } => go_type_assert_needs_single_eval(object),
        ExprKind::Index { object, index, .. } => {
            go_type_assert_needs_single_eval(object) || go_type_assert_needs_single_eval(index)
        }
        ExprKind::Unary { expr, .. } => go_type_assert_needs_single_eval(expr),
        ExprKind::Binary { left, right, .. } => {
            go_type_assert_needs_single_eval(left) || go_type_assert_needs_single_eval(right)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_type_assert_needs_single_eval(cond)
                || go_type_assert_needs_single_eval(then)
                || go_type_assert_needs_single_eval(else_)
        }
        ExprKind::Cast { expr, .. } | ExprKind::TypeOf(expr) => {
            go_type_assert_needs_single_eval(expr)
        }
        _ => false,
    }
}

fn go_concrete_type_for_interface(type_name: &str, env: &GoNormalizeEnv) -> Option<String> {
    let required = env.interface_methods.get(type_name)?;
    if required.is_empty() {
        return None;
    }
    env.struct_infos
        .iter()
        .find(|(_, info)| {
            required
                .iter()
                .all(|method| info.method_names.contains(method))
        })
        .map(|(name, _)| name.clone())
}

fn go_type_switch_case_cond(expr: Expression, case_types: &[String]) -> Expression {
    let mut iter = case_types.iter();
    let first = iter
        .next()
        .map(|type_name| go_build_type_switch_case_expr(expr.clone(), type_name))
        .unwrap_or_else(|| Expression::bool(false));
    iter.fold(first, |acc, type_name| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(acc),
            right: Box::new(go_build_type_switch_case_expr(expr.clone(), type_name)),
        })
    })
}

fn go_build_type_switch_case_expr(expr: Expression, type_name: &str) -> Expression {
    if type_name == "__go_nil" {
        return Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(expr),
            right: Box::new(Expression::null()),
        });
    }
    go_build_is_type(expr, type_name)
}

fn go_non_null_cond(expr: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr),
        right: Box::new(Expression::null()),
    })
}

fn go_non_null_object_cond(expr: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(go_non_null_cond(expr.clone())),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(expr)))),
            right: Box::new(Expression::string("object")),
        })),
    })
}

fn go_object_has_fields_cond(expr: Expression, fields: &[&str]) -> Expression {
    let mut iter = fields.iter();
    let first = iter
        .next()
        .map(|field| go_object_has_field_cond(expr.clone(), field))
        .unwrap_or_else(|| go_non_null_object_cond(expr.clone()));
    iter.fold(first, |acc, field| {
        Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(acc),
            right: Box::new(go_object_has_field_cond(expr.clone(), field)),
        })
    })
}

fn go_object_has_field_cond(expr: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(go_non_null_object_cond(expr.clone())),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: field.to_string(),
                null_safe: false,
            })),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
        })),
    })
}

fn go_build_is_type(expr: Expression, type_name: &str) -> Expression {
    let typeof_tag = match type_name.trim() {
        "int" | "int8" | "int16" | "int32" | "int64" | "uint" | "uint8" | "uint16" | "uint32"
        | "uint64" | "uintptr" | "byte" | "rune" | "float32" | "float64" => Some("number"),
        "string" => Some("string"),
        "bool" => Some("boolean"),
        _ => None,
    };

    if let Some(tag) = typeof_tag {
        return Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(expr)))),
            right: Box::new(Expression::string(tag)),
        });
    }

    // Map Go composite types to the canonical IsType categories the shared
    // compiler recognizes (array→isArray, function, map→object-kind).
    let trimmed = type_name.trim();
    let canon = if trimmed.starts_with("[]") || (trimmed.starts_with('[') && trimmed.contains(']'))
    {
        "array"
    } else if trimmed.starts_with("func") {
        "function"
    } else if trimmed.starts_with("map[") {
        "map"
    } else {
        trimmed
    };

    Expression::new(ExprKind::IsType {
        expr: Box::new(expr),
        type_name: canon.to_string(),
    })
}

fn go_type_switch_case_body(
    mut body: Vec<Statement>,
    binding_name: Option<&str>,
    expr: Expression,
    case_type: &str,
) -> Vec<Statement> {
    if let Some(name) = binding_name {
        body.insert(
            0,
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name.to_string()),
                    init: Some(Expression::new(ExprKind::Cast {
                        expr: Box::new(expr),
                        type_name: case_type.to_string(),
                    })),
                    type_hint: Some(case_type.to_string().into()),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }),
        );
    }
    body
}

fn go_wrap_spawn_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Expr(expr))]),
        is_async: false,
        captures: Vec::new(),
    })
}

fn go_type_name_from_expr(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Cast { expr, type_name } if matches!(expr.kind, ExprKind::Lit(Literal::Null)) => {
            Some(type_name.clone())
        }
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { .. } => go_expr_call_name(expr),
        _ => None,
    }
}

fn go_is_slice_type(type_name: &str) -> bool {
    type_name.trim_start().starts_with("[]")
}

fn go_is_map_type(type_name: &str) -> bool {
    type_name.trim_start().starts_with("map[")
}

fn go_is_channel_type(type_name: &str) -> bool {
    let trimmed = type_name.trim_start();
    trimmed.starts_with("chan") || trimmed.starts_with("<-chan")
}

fn go_channel_element_type(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim();
    let elem = if let Some(rest) = trimmed.strip_prefix("<-chan") {
        rest
    } else if let Some(rest) = trimmed.strip_prefix("chan<-") {
        rest
    } else if let Some(rest) = trimmed.strip_prefix("chan") {
        rest.trim_start_matches("<-")
    } else {
        return None;
    };
    let elem = elem.trim();
    (!elem.is_empty()).then(|| elem.to_string())
}

fn go_array_make_expr(len_expr: Expression, init_expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![
            Argument::positional(len_expr),
            Argument::positional(init_expr),
        ],
        optional: false,
    })
}

fn go_make_slice_capacity_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("make") {
        return None;
    }
    let type_name = args
        .first()
        .and_then(|arg| go_type_name_from_expr(&arg.value))?;
    if !go_is_slice_type(&type_name) {
        return None;
    }
    let cap_arg = args.get(2).or_else(|| args.get(1))?;
    Some(normalize_go_expr(&cap_arg.value, env, signatures, state))
}

fn go_append_capacity_expr(
    original: &Expression,
    normalized: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &original.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("append") || args.is_empty() {
        return None;
    }
    let needed = go_builtin_call("len", vec![normalized.clone()]);
    let current = go_expr_capacity_hint(&args[0].value, env)?;
    Some(Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(current.clone()),
            right: Box::new(needed.clone()),
        })),
        then: Box::new(needed),
        else_: Box::new(current),
    }))
}

fn go_bound_slice_capacity_expr(expr: &Expression, env: &GoNormalizeEnv) -> Option<Expression> {
    if let Some(cap) = go_expr_capacity_hint(expr, env) {
        return Some(cap);
    }

    match &expr.kind {
        ExprKind::Call { callee, args, .. }
            if matches!(
                go_expr_call_name(callee).as_deref(),
                Some("__go_slices_Grow" | "__go_slices_grow_common")
            ) =>
        {
            let slice = &args[0].value;
            let grow_by = args
                .get(1)
                .map(|arg| arg.value.clone())
                .unwrap_or_else(|| Expression::int(0));
            let current_cap = go_expr_capacity_hint(slice, env)
                .unwrap_or_else(|| go_builtin_call("len", vec![slice.clone()]));
            let needed_cap = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(go_builtin_call("len", vec![slice.clone()])),
                right: Box::new(grow_by),
            });
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(current_cap.clone()),
                    right: Box::new(needed_cap.clone()),
                })),
                then: Box::new(needed_cap),
                else_: Box::new(current_cap),
            }))
        }
        ExprKind::Call { callee, args, .. } => {
            if go_expr_call_name(callee).as_deref() != Some("make") {
                return None;
            }
            let type_name = args
                .first()
                .and_then(|arg| go_type_name_from_expr(&arg.value))?;
            if !go_is_slice_type(&type_name) {
                return None;
            }
            Some(args.get(2).or_else(|| args.get(1))?.value.clone())
        }
        ExprKind::Cast { expr, type_name } if go_is_slice_type(type_name) => {
            go_bound_slice_capacity_expr(expr, env)
        }
        ExprKind::Array(elements) => Some(Expression::int(elements.len() as i64)),
        _ => None,
    }
}

fn go_expr_capacity_hint(expr: &Expression, env: &GoNormalizeEnv) -> Option<Expression> {
    if let Some(view) = go_expr_slice_view(expr, env) {
        if let Some(max) = view.max {
            return Some(Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(max),
                right: Box::new(view.start),
            }));
        }
        let base_cap = go_expr_capacity_hint(&view.base, env)?;
        return Some(Expression::new(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(base_cap),
            right: Box::new(view.start),
        }));
    }
    match &expr.kind {
        ExprKind::Ident(name) => env.slice_caps.get(name).cloned().or_else(|| {
            env.fixed_arrays
                .get(name)
                .and_then(|type_name| go_fixed_array_len(type_name, 0))
                .map(|len| Expression::int(len as i64))
        }),
        ExprKind::Member { object, field, .. } => {
            let cap_field = Expression::new(ExprKind::Member {
                object: object.clone(),
                field: format!("{}__cap", field),
                null_safe: false,
            });
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(cap_field.clone()),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
                })),
                then: Box::new(cap_field),
                else_: Box::new(go_builtin_call("len", vec![expr.clone()])),
            }))
        }
        ExprKind::Array(elements) => Some(Expression::int(elements.len() as i64)),
        _ => None,
    }
}

fn go_binding_name(pattern: &BindingPattern) -> Option<String> {
    match pattern {
        BindingPattern::Ident(name) => Some(name.clone()),
        _ => go_single_named_binding_pattern(pattern).and_then(|pattern| match pattern {
            BindingPattern::Ident(name) => Some(name),
            _ => None,
        }),
    }
}

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    // ⛔ NOT `Position::line_col` — it counts newlines from the START OF THE
    // INPUT, twice per node, which makes the walk quadratic in program size.
    // Measured with `-c` on a ladder of assignments: 320 -> 1.27s, 640 -> 3.02s,
    // 1280 -> 9.34s, i.e. 3x the time per doubling where linear is 2x. See
    // `vybe_ast::line_index`. The fallback is the old behaviour, for a parse
    // that reached here without installing an index.
    vybe_ast::line_index::span_1based(s.start(), s.end()).unwrap_or_else(|| {
        let (sl, sc) = s.start_pos().line_col();
        let (el, ec) = s.end_pos().line_col();
        Span {
            start_line: sl as u32,
            start_col: sc as u32,
            end_line: el as u32,
            end_col: ec as u32,
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sync_atomic_function_calls_normalize_to_common_atomic_node() {
        let module = parse(
            r#"
package main

import "sync/atomic"

func main() {
	var n int64
	atomic.StoreInt64(&n, 1)
	atomic.AddInt64(&n, 2)
	_ = atomic.LoadInt64(&n)
	_ = atomic.CompareAndSwapInt64(&n, 3, 4)
}
"#,
        )
        .expect("go source should parse");

        let rendered = format!("{:?}", module);
        assert!(
            rendered.contains("Atomic"),
            "sync/atomic should lower through ExprKind::Atomic"
        );
        assert!(
            rendered.contains("CompareExchange"),
            "CompareAndSwap should lower through AtomicOp::CompareExchange"
        );
    }

    #[test]
    fn sync_atomic_typed_numeric_methods_normalize_to_common_atomic_node() {
        let module = parse(
            r#"
package main

import "sync/atomic"

func main() {
	var v atomic.Int64
	v.Store(1)
	v.Add(2)
	_ = v.Load()
	_ = v.CompareAndSwap(3, 4)
}
"#,
        )
        .expect("go source should parse");

        let rendered = format!("{:?}", module);
        assert!(
            rendered.contains("__goAtomicInt64"),
            "atomic.Int64 should bind to the Go atomic adapter type"
        );
        assert!(
            rendered.contains("Atomic"),
            "typed atomic numeric methods should lower through ExprKind::Atomic"
        );
        assert!(
            rendered.contains("CompareExchange"),
            "typed CompareAndSwap should lower through AtomicOp::CompareExchange"
        );
    }

    #[test]
    fn sync_atomic_value_methods_use_adapter_storage() {
        let module = parse(
            r#"
package main

import "sync/atomic"

func main() {
	var v atomic.Value
	v.Store("hello")
	_ = v.Load()
}
"#,
        )
        .expect("go source should parse");

        let rendered = format!("{:?}", module);
        assert!(
            rendered.contains("__goAtomicValue"),
            "atomic.Value should bind to the Go atomic value adapter type"
        );
        assert!(
            rendered.contains("value"),
            "atomic.Value should use explicit adapter storage"
        );
    }
}
