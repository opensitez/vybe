//! Bash → common AST.
//!
//! Bash has no expression grammar of its own: a program is a list of commands
//! whose only value is an exit status, and words are concatenations of parts.
//! The walker therefore lowers along two axes:
//!
//! * **Commands** become statements. A simple command is a `Call` of the
//!   command name; builtins that are control flow in bash (`return`, `exit`,
//!   `break`, `continue`, `local`, `declare`, …) become the matching AST nodes.
//! * **Conditions** are booleans. A `[ ]`/`[[ ]]` test or `(( ))` lowers to a
//!   boolean expression directly; any other command used as a condition is
//!   wrapped as `!status` (bash: status 0 is success, JS: 0 is falsy), so
//!   `if cmd`, `cmd && other`, `while ! cmd` keep bash polarity.
//!
//! Every word is a STRING. Arithmetic contexts coerce operands with unary `+`
//! (an unset or empty variable is 0), exactly as bash does.
//!
//! The shell state bash keeps implicitly is kept in ordinary variables the
//! walker declares at the top of the module:
//! * `__bash_args` — the positional parameters (`$1`, `$#`, `"$@"`); every
//!   function declares its own parameter with the same name.
//! * `__bash_status` — `$?`, assigned after every statement-level command.
//! * `__bash_stdin` — the text a redirection (`<<<`, `<<`, `< file`,
//!   `< <(…)`) or a pipeline made available to `read`/`mapfile`.
//! * `BASH_REMATCH` — set by `[[ s =~ re ]]`.
//!
//! Output capture (`$(…)`, `> file`, `>/dev/null`, pipelines) uses the shared
//! output-buffer operations bound in the profile as `__bash_ob_*`; file
//! operations use `__bash_write_file` & co. What has no shared operation yet
//! (process substitution as an argument, dynamic glob patterns, `-x`/`-nt`
//! file tests, `eval`, `cd`, …) is lowered to a call of a `__bash_*` name so
//! the AST states the operation instead of dropping it.

use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};

use pest::Parser;
use pest::iterators::Pair;
use vybe_ast::*;

use crate::{Rule, WashmParser};

/// Variables the walker declares before the program's own statements.
const PRELUDE_DECLS: usize = 28;
const BASH_ARGS: &str = "__bash_args";
const BASH_FUNCNAME: &str = "__bash_funcname";

// ── entry ────────────────────────────────────────────────────────────────────

/// Parse Bash source into the common AST.
pub fn parse(source: &str) -> Result<Module, String> {
    let (stripped, heredocs) = extract_heredocs(source);
    let _line_index = vybe_ast::line_index::LineIndex::install(&stripped);
    let pairs =
        WashmParser::parse(Rule::program, &stripped).map_err(|e| format!("Parse error: {e}"))?;

    let mut w = Walker {
        heredocs: heredocs.into(),
        counter: 0,
        functions: HashMap::new(),
        function_bodies: HashMap::new(),
        function_commands: HashMap::new(),
        variables: [
            "HOME".to_string(),
            "PWD".to_string(),
            "OLDPWD".to_string(),
            "BASH_SUBSHELL".to_string(),
        ]
        .into_iter()
        .collect(),
        variable_values: HashMap::new(),
        array_values: HashMap::new(),
        indexed_arrays: HashSet::new(),
        assoc_arrays: HashSet::new(),
        readonly_vars: HashSet::new(),
        integer_vars: HashSet::new(),
        lowercase_vars: HashSet::new(),
        uppercase_vars: HashSet::new(),
        exported_vars: HashSet::new(),
        exported_functions: HashSet::new(),
        traced_functions: HashSet::new(),
        namerefs: HashMap::new(),
        aliases: HashMap::new(),
        positional_values: Vec::new(),
        brace_expansion_enabled: true,
        nullglob_enabled: false,
        dotglob_enabled: false,
        failglob_enabled: false,
        globstar_enabled: false,
        expand_aliases_enabled: false,
        nocaseglob_enabled: false,
        nocasematch_enabled: false,
        patsub_replacement_enabled: true,
        lastpipe_enabled: false,
        shell_flags: "Bh".chars().collect(),
        subshell_depth: 0,
        arith_stack: Vec::new(),
        dynamic_arith: false,
        function_depth: 0,
        inline_call_context: false,
        inline_stack: Vec::new(),
        suppress_function_inlining: false,
        hoisted_functions: Vec::new(),
        exit_trap_body: None,
    };
    let mut body = Vec::new();
    for (name, init) in [
        (BASH_ARGS, array(Vec::new())),
        ("__bash_status", int(0)),
        ("__bash_ret", undefined()),
        ("__bash_stdin", lit("")),
        ("__bash_fds", object(Vec::new())),
        ("__bash_stdout_null", Expression::bool(false)),
        ("__bash_stdout_to_stderr", Expression::bool(false)),
        ("__bash_stderr_to_stdout", Expression::bool(false)),
        ("__bash_stderr_null", Expression::bool(false)),
        ("__bash_abort_line", Expression::bool(false)),
        ("__bash_line", int(0)),
        ("BASH_REMATCH", array(Vec::new())),
        ("PIPESTATUS", array(Vec::new())),
        (BASH_FUNCNAME, array(Vec::new())),
        ("__bash_had", Expression::bool(false)),
        ("__bash_fields", array(Vec::new())),
        ("__bash_last_arg", lit("")),
        ("__bash_pid", int(1)),
        ("BASH_SUBSHELL", int(0)),
        ("__bash_last_bg_pid", int(0)),
        ("__bash_job_seq", int(1)),
        ("__bash_jobs", object(Vec::new())),
        ("__bash_coprocs", object(Vec::new())),
        ("__bash_flags", lit("Bh")),
        ("TIMEFORMAT", lit("real 0.00\nuser 0.00\nsys 0.00")),
        ("HOME", bash_env_value_expr("HOME", lit(""))),
        ("PWD", call_named("__bash_getcwd", vec![])),
        ("OLDPWD", undefined()),
    ] {
        body.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![declarator(name, Some(init))],
            kind: VarDeclKind::Var,
        }));
    }
    debug_assert_eq!(body.len(), PRELUDE_DECLS);
    for pair in pairs {
        if pair.as_rule() == Rule::program {
            for inner in pair.into_inner() {
                if inner.as_rule() == Rule::list {
                    body.extend(w.walk_list(inner)?);
                }
            }
        }
    }
    let predeclared = bash_user_variable_predecls(&w.variables);
    let mut insert_at = PRELUDE_DECLS;
    if !predeclared.is_empty() {
        body.insert(
            insert_at,
            Statement::new(StmtKind::VarDecl {
                declarations: predeclared,
                kind: VarDeclKind::FunctionScoped,
            }),
        );
        insert_at += 1;
    }
    if !w.hoisted_functions.is_empty() {
        body.splice(insert_at..insert_at, w.hoisted_functions);
    }
    Ok(Module {
        canon: Default::default(),
        name: "main".into(),
        language: Lang::Unknown,
        body,
        imports: Vec::new(),
        directives: Directives {
            spread_arguments: Some(SpreadArguments::Positional),
            ..Default::default()
        },
    })
}

/// One here-document body cut out of the source by [`extract_heredocs`].
#[derive(Clone)]
struct HereDoc {
    body: String,
    /// The delimiter was unquoted, so `$var`, `$(…)` and `` `…` `` expand.
    expand: bool,
    eof_warning: Option<String>,
}

/// Remove every here-document body from `source`, in order of appearance.
///
/// The grammar only sees the `<<DELIM` marker; the body runs from the line
/// after the marker to the line holding only the delimiter (leading tabs
/// stripped for `<<-`). A quoted or backslashed delimiter disables expansion.
fn extract_heredocs(source: &str) -> (String, Vec<HereDoc>) {
    let lines: Vec<&str> = source.split('\n').collect();
    let mut out: Vec<String> = Vec::new();
    let mut docs = Vec::new();
    let mut i = 0;
    while i < lines.len() {
        let line = lines[i];
        out.push(line.to_string());
        let markers = heredoc_markers(line);
        i += 1;
        for (delim, strip, expand) in markers {
            let mut body = String::new();
            let mut terminated = false;
            while i < lines.len() {
                let raw = lines[i];
                let candidate = if strip {
                    raw.trim_start_matches('\t')
                } else {
                    raw
                };
                i += 1;
                if candidate == delim {
                    terminated = true;
                    break;
                }
                body.push_str(candidate);
                body.push('\n');
            }
            let eof_warning = (!terminated).then(|| {
                format!(
                    "bash: warning: here-document delimited by end-of-file (wanted `{delim}')\n"
                )
            });
            docs.push(HereDoc {
                body,
                expand,
                eof_warning,
            });
        }
    }
    (out.join("\n"), docs)
}

/// The `<<[-]DELIM` markers on one line, outside quotes, left to right.
fn heredoc_markers(line: &str) -> Vec<(String, bool, bool)> {
    let chars: Vec<char> = line.chars().collect();
    let mut found = Vec::new();
    let (mut in_single, mut in_double) = (false, false);
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '\\' if !in_single => {
                i += 2;
                continue;
            }
            '\'' if !in_double => in_single = !in_single,
            '"' if !in_single => in_double = !in_double,
            '#' if !in_single && !in_double && (i == 0 || chars[i - 1].is_whitespace()) => break,
            '<' if !in_single && !in_double && i + 1 < chars.len() && chars[i + 1] == '<' => {
                // `<<<` is a here-string.
                if i + 2 < chars.len() && chars[i + 2] == '<' {
                    i += 3;
                    continue;
                }
                i += 2;
                let strip = i < chars.len() && chars[i] == '-';
                if strip {
                    i += 1;
                }
                while i < chars.len() && (chars[i] == ' ' || chars[i] == '\t') {
                    i += 1;
                }
                let mut delim = String::new();
                let mut expand = true;
                while i < chars.len() {
                    let d = chars[i];
                    match d {
                        '\'' | '"' => {
                            expand = false;
                            i += 1;
                            while i < chars.len() && chars[i] != d {
                                delim.push(chars[i]);
                                i += 1;
                            }
                            i += 1;
                        }
                        '\\' => {
                            expand = false;
                            if i + 1 < chars.len() {
                                delim.push(chars[i + 1]);
                            }
                            i += 2;
                        }
                        ' ' | '\t' | ';' | '|' | '&' | ')' | '(' | '<' | '>' => break,
                        _ => {
                            delim.push(d);
                            i += 1;
                        }
                    }
                }
                if !delim.is_empty() {
                    found.push((delim, strip, expand));
                }
                continue;
            }
            _ => {}
        }
        i += 1;
    }
    found
}

// ── small constructors ───────────────────────────────────────────────────────

fn ident(name: &str) -> Expression {
    Expression::ident(name)
}
fn lit(s: &str) -> Expression {
    Expression::string(s)
}
fn int(n: i64) -> Expression {
    Expression::int(n)
}
fn bigint(n: i64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::BigInt(n)))
}
fn undefined() -> Expression {
    Expression::new(ExprKind::Lit(Literal::Undefined))
}
fn null() -> Expression {
    Expression::null()
}
fn binary(op: BinOp, l: Expression, r: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(l),
        right: Box::new(r),
    })
}
fn unary(op: UnaryOp, e: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op,
        expr: Box::new(e),
    })
}
fn ternary(c: Expression, t: Expression, e: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(c),
        then: Box::new(t),
        else_: Box::new(e),
    })
}
fn member(obj: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(obj),
        field: field.to_string(),
        null_safe: false,
    })
}
fn index(obj: Expression, idx: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(obj),
        index: Box::new(idx),
        null_safe: false,
    })
}
fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}
fn call_named(name: &str, args: Vec<Expression>) -> Expression {
    call(ident(name), args)
}
fn call_named_args(name: &str, args: Vec<Argument>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(ident(name)),
        args,
        optional: false,
    })
}
fn bash_sprintf_array(fmt: Expression, args: Vec<ShellArg>) -> Expression {
    call_named("__bash_sprintf_array", vec![fmt, shell_args_array(args)])
}
fn object(props: Vec<(&str, Expression)>) -> Expression {
    Expression::new(ExprKind::Object(
        props
            .into_iter()
            .map(|(key, value)| ObjectProperty::KeyValue {
                key: lit(key),
                value,
            })
            .collect(),
    ))
}
fn method(obj: Expression, name: &str, args: Vec<Expression>) -> Expression {
    call(member(obj, name), args)
}
fn bash_path_expr(path: Expression) -> Expression {
    bash_path_with_cwd(path, ident("PWD"))
}
fn bash_path_with_cwd(path: Expression, cwd: Expression) -> Expression {
    ternary(
        method(path.clone(), "startsWith", vec![lit("/")]),
        path.clone(),
        parts_to_expr(vec![
            Part::Expr(cwd),
            Part::Text("/".into()),
            Part::Expr(path),
        ]),
    )
}
fn assign_expr(target: Expression, value: Expression) -> Expression {
    Expression::new(ExprKind::Assign {
        target: Box::new(target),
        value: Box::new(value),
    })
}
fn assign_stmt(target: Expression, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![target],
        value,
        by_ref: false,
    })
}
fn sequence(items: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Sequence(items))
}
fn array(items: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        items
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
fn array_spread(items: Vec<(Expression, bool)>) -> Expression {
    Expression::new(ExprKind::Array(
        items
            .into_iter()
            .map(|(value, spread)| ArrayElement {
                key: None,
                value,
                spread,
                by_ref: false,
            })
            .collect(),
    ))
}
fn shell_args_array(args: Vec<ShellArg>) -> Expression {
    if args.len() == 1 && args[0].spread {
        return args.into_iter().next().unwrap().value;
    }
    Expression::new(ExprKind::Array(
        args.into_iter()
            .map(|arg| ArrayElement {
                key: None,
                value: arg.value,
                spread: arg.spread,
                by_ref: false,
            })
            .collect(),
    ))
}
fn filter_non_empty_array_expr(items: Expression) -> Expression {
    method(
        items,
        "filter",
        vec![lambda_expr(
            vec![param_named("__bash_array_item")],
            binary(BinOp::StrictNotEq, ident("__bash_array_item"), lit("")),
        )],
    )
}
/// `new RegExp(source, flags)`.
fn regexp(source: Expression, flags: &str) -> Expression {
    object(vec![("source", source), ("flags", lit(flags))])
}
fn lambda_block(body: Vec<Statement>) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    })
}
fn lambda_block_with_params(params: Vec<Param>, body: Vec<Statement>) -> Expression {
    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    })
}
fn lambda_expr(params: Vec<Param>, body: Expression) -> Expression {
    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Expr(Box::new(body)),
        is_async: false,
        captures: Vec::new(),
    })
}
fn case_toggle_lambda() -> Expression {
    lambda_expr(
        vec![param_named("__bash_ch")],
        ternary(
            binary(
                BinOp::StrictEq,
                ident("__bash_ch"),
                method(ident("__bash_ch"), "toUpperCase", vec![]),
            ),
            method(ident("__bash_ch"), "toLowerCase", vec![]),
            method(ident("__bash_ch"), "toUpperCase", vec![]),
        ),
    )
}
/// `(() => { body })()` — a block evaluated as an expression.
fn iife(body: Vec<Statement>) -> Expression {
    call(lambda_block(body), Vec::new())
}
fn iife_with_args(body: Vec<Statement>, params: Vec<Param>, args: Vec<Expression>) -> Expression {
    call(lambda_block_with_params(params, body), args)
}
fn shell_state_params_args() -> (Vec<Param>, Vec<Expression>) {
    let names = [
        "BASH_SUBSHELL",
        BASH_FUNCNAME,
        "PIPESTATUS",
        "TIMEFORMAT",
        "__bash_fds",
        "__bash_ret",
        "__bash_status",
        "__bash_stdin",
        "__bash_abort_line",
        "__bash_jobs",
        "__bash_coprocs",
        "__bash_job_seq",
        "__bash_stderr_null",
        "__bash_stderr_to_stdout",
        "__bash_stdout_null",
        "__bash_stdout_to_stderr",
    ];
    (
        names.iter().map(|name| param_named(name)).collect(),
        names.iter().map(|name| ident(name)).collect(),
    )
}
fn shell_state_iife(body: Vec<Statement>) -> Expression {
    let (params, args) = shell_state_params_args();
    iife_with_args(body, params, args)
}
fn bash_input_lines(input: Expression) -> Expression {
    let value = "__bash_line_input";
    let lines = "__bash_line_items";
    iife_with_args(
        vec![
            Statement::new(StmtKind::If {
                cond: binary(BinOp::StrictEq, ident(value), lit("")),
                then_body: vec![Statement::new(StmtKind::Return(Some(array(Vec::new()))))],
                elifs: Vec::new(),
                else_body: None,
            }),
            let_stmt(lines, method(ident(value), "split", vec![lit("\n")])),
            Statement::new(StmtKind::If {
                cond: method(ident(value), "endsWith", vec![lit("\n")]),
                then_body: vec![expr_stmt(method(ident(lines), "pop", vec![]))],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::Return(Some(ident(lines)))),
        ],
        vec![param_named(value)],
        vec![input],
    )
}
fn bash_function_params() -> Vec<Param> {
    let (state_params, _) = shell_state_params_args();
    let mut params = vec![args_param()];
    params.extend(state_params);
    params
}
fn funcname_push_stmts(name: &str) -> Vec<Statement> {
    let mut out = Vec::new();
    for i in (1..32).rev() {
        out.push(assign_stmt(
            index(ident(BASH_FUNCNAME), int(i)),
            index(ident(BASH_FUNCNAME), int(i - 1)),
        ));
    }
    out.push(assign_stmt(index(ident(BASH_FUNCNAME), int(0)), lit(name)));
    out
}
fn funcname_pop_exprs() -> Vec<Expression> {
    let mut out = Vec::new();
    for i in 0..31 {
        out.push(assign_expr(
            index(ident(BASH_FUNCNAME), int(i)),
            index(ident(BASH_FUNCNAME), int(i + 1)),
        ));
    }
    out.push(assign_expr(index(ident(BASH_FUNCNAME), int(31)), undefined()));
    out
}
fn declarator(name: &str, init: Option<Expression>) -> VarDeclarator {
    VarDeclarator {
        pattern: BindingPattern::Ident(name.to_string()),
        type_hint: None,
        init,
        array_bounds: None,
        with_events: false,
    }
}
fn let_stmt(name: &str, init: Expression) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![declarator(name, Some(init))],
        kind: VarDeclKind::Let,
    })
}
fn expr_stmt(e: Expression) -> Statement {
    Statement::new(StmtKind::Expr(e))
}

fn stmts_to_exprs(stmts: &[Statement]) -> Option<Vec<Expression>> {
    let mut out = Vec::new();
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::Expr(expr) => out.push(expr.clone()),
            StmtKind::Assign {
                targets,
                value,
                by_ref: false,
            } if targets.len() == 1 => out.push(assign_expr(targets[0].clone(), value.clone())),
            StmtKind::VarDecl { declarations, .. } => {
                for declarator in declarations {
                    if let BindingPattern::Ident(name) = &declarator.pattern {
                        out.push(assign_expr(
                            ident(name),
                            declarator.init.clone().unwrap_or_else(undefined),
                        ));
                    } else {
                        return None;
                    }
                }
            }
            StmtKind::Block(body) => out.extend(stmts_to_exprs(body)?),
            _ => return None,
        }
    }
    Some(out)
}

/// The positional-parameter array every bash function declares: `$1`, `$#`, `$@` read it.
fn args_param() -> Param {
    param_named(BASH_ARGS)
}
fn param_named(name: &str) -> Param {
    Param {
        name: name.into(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}
/// Numeric value of a word in an arithmetic context: unset/empty is 0.
fn to_number(e: Expression) -> Expression {
    match &e.kind {
        ExprKind::Lit(Literal::Int(_)) | ExprKind::Lit(Literal::Float(_)) => e,
        ExprKind::Lit(Literal::BigInt(_)) => e,
        ExprKind::Lit(Literal::Str(s)) => match parse_bash_integer(s) {
            Some(n) => int(n),
            None => unary(UnaryOp::Pos, e),
        },
        _ => iife(vec![
            let_stmt("__bash_num", e),
            Statement::new(StmtKind::Return(Some(ternary(
                ident("__bash_num"),
                unary(UnaryOp::Pos, ident("__bash_num")),
                int(0),
            )))),
        ]),
    }
}
fn arith_i64(e: Expression) -> Expression {
    call_named(
        "__bash_i64_wrap",
        vec![int(64), call_named("__bash_bigint", vec![arith_value(e)])],
    )
}
fn arith_value(e: Expression) -> Expression {
    iife_with_args(
        vec![Statement::new(StmtKind::Return(Some(ternary(
            binary(
                BinOp::Or,
                binary(BinOp::StrictEq, ident("__bash_arith_value"), undefined()),
                binary(BinOp::StrictEq, ident("__bash_arith_value"), lit("")),
            ),
            lit("0"),
            ident("__bash_arith_value"),
        ))))],
        vec![param_named("__bash_arith_value")],
        vec![e],
    )
}
fn arith_wrap(e: Expression) -> Expression {
    call_named("__bash_i64_wrap", vec![int(64), e])
}
fn arith_string(e: Expression) -> Expression {
    call_named("__bash_i64_string", vec![e])
}
fn arith_index(e: Expression) -> Expression {
    if matches!(&e.kind, ExprKind::Lit(Literal::Int(_))) {
        return e;
    }
    to_number(arith_string(e))
}
fn arith_bin(name: &str, left: Expression, right: Expression) -> Expression {
    arith_wrap(call_named(name, vec![arith_i64(left), arith_i64(right)]))
}
fn arith_unary(name: &str, value: Expression) -> Expression {
    arith_wrap(call_named(name, vec![arith_i64(value)]))
}
fn arith_cmp(name: &str, left: Expression, right: Expression) -> Expression {
    call_named(name, vec![arith_i64(left), arith_i64(right)])
}
fn arith_cmp_i64(name: &str, left: Expression, right: Expression) -> Expression {
    arith_bool_to_i64(arith_cmp(name, left, right))
}
fn arith_bool(e: Expression) -> Expression {
    arith_cmp("__bash_i64_ne", e, bigint(0))
}
fn arith_bool_to_i64(e: Expression) -> Expression {
    ternary(e, bigint(1), bigint(0))
}
/// `$?` from a command's value: `false` is 1, a number is itself, anything
/// else (no return value, `true`, a string) is success.
fn status_of_value(v: Expression) -> Expression {
    ternary(
        binary(BinOp::StrictEq, v.clone(), Expression::bool(false)),
        int(1),
        ternary(
            binary(
                BinOp::StrictEq,
                Expression::new(ExprKind::TypeOf(Box::new(v.clone()))),
                lit("number"),
            ),
            v,
            int(0),
        ),
    )
}
/// Capture what `body` prints. Command substitution runs in a subshell, so it
/// inherits the current positional array but mutations stay inside the wrapper.
fn capture(body: Vec<Statement>) -> Expression {
    capture_with_args(body, vec![args_param()], vec![ident(BASH_ARGS)])
}

fn capture_with_args(
    body: Vec<Statement>,
    params: Vec<Param>,
    args: Vec<Expression>,
) -> Expression {
    let mut body = body;
    rewrite_exit_to_return(&mut body);
    body.push(Statement::new(StmtKind::Return(Some(ident(
        "__bash_status",
    )))));
    sequence(vec![
        call_named("__bash_ob_start", vec![]),
        assign_expr(ident("__bash_status"), iife_with_args(body, params, args)),
        call_named("__bash_ob_get_clean", vec![]),
    ])
}

fn current_shell_capture(body: Vec<Statement>, rewrite_exit: bool) -> Expression {
    let mut body = body;
    if rewrite_exit {
        rewrite_exit_to_subshell_throw(&mut body);
    }
    body.push(Statement::new(StmtKind::Return(Some(ident(
        "__bash_status",
    )))));
    sequence(vec![
        call_named("__bash_ob_start", vec![]),
        assign_expr(ident("__bash_status"), iife(body)),
        call_named("__bash_ob_get_clean", vec![]),
    ])
}

fn array_join(items: Expression, sep: Expression) -> Expression {
    call_named("__bash_join", vec![items, sep])
}

fn ifs_join_sep() -> Expression {
    ternary(
        binary(BinOp::StrictEq, ident("IFS"), undefined()),
        lit(" "),
        method(param_value(ident("IFS")), "charAt", vec![int(0)]),
    )
}

fn split_bash_words(value: Expression) -> Expression {
    let raw = "__bash_split_value";
    let trimmed = "__bash_split_trimmed";
    let sep = "__bash_split_sep";
    let parts = "__bash_split_parts";
    iife(vec![
        let_stmt(raw, call_named("__bash_string", vec![param_value(value)])),
        Statement::new(StmtKind::If {
            cond: binary(BinOp::StrictEq, ident(raw), lit("")),
            then_body: vec![Statement::new(StmtKind::Return(Some(array(Vec::new()))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: binary(BinOp::StrictEq, ident("IFS"), lit("")),
            then_body: vec![Statement::new(StmtKind::Return(Some(array(vec![ident(
                raw,
            )]))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        let_stmt(trimmed, method(ident(raw), "trim", vec![])),
        Statement::new(StmtKind::If {
            cond: binary(BinOp::StrictEq, ident(trimmed), lit("")),
            then_body: vec![Statement::new(StmtKind::Return(Some(array(Vec::new()))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Or,
                binary(BinOp::StrictEq, ident("IFS"), undefined()),
                binary(BinOp::StrictEq, ident("IFS"), lit(" \t\n")),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(call_named(
                "__bash_re_split",
                vec![ident(trimmed), lit("[ \t\r\n]+")],
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        let_stmt(
            sep,
            method(param_value(ident("IFS")), "charAt", vec![int(0)]),
        ),
        let_stmt(parts, method(ident(raw), "split", vec![ident(sep)])),
        Statement::new(StmtKind::If {
            cond: method(ident(raw), "endsWith", vec![ident(sep)]),
            then_body: vec![expr_stmt(method(ident(parts), "pop", vec![]))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(ident(parts)))),
    ])
}
fn regex_test_expr(pattern: Expression, subject: Expression) -> Expression {
    call_named("__bash_regex_test", vec![pattern, subject])
}
fn regex_exec_expr(pattern: Expression, subject: Expression) -> Expression {
    call_named("__bash_regex_exec", vec![pattern, subject])
}
fn regex_replace_expr(
    target: Expression,
    source: &str,
    flags: &str,
    replacement: Expression,
) -> Expression {
    call_named(
        "__bash_regex_replace",
        vec![target, regexp(lit(source), flags), replacement],
    )
}
fn bash_quote_expr(value: Expression) -> Expression {
    let quoted = parts_to_expr(vec![
        Part::Text("'".into()),
        Part::Expr(regex_replace_expr(value.clone(), "'", "g", lit("'\\''"))),
        Part::Text("'".into()),
    ]);
    ternary(binary(BinOp::StrictEq, value, lit("")), lit("''"), quoted)
}

fn positional_slice_start(offset: Expression) -> Expression {
    match &offset.kind {
        ExprKind::Lit(Literal::Int(0)) => int(0),
        ExprKind::Lit(Literal::Int(n)) if *n > 0 => int(n - 1),
        ExprKind::Lit(Literal::Int(n)) if *n < 0 => {
            binary(BinOp::Add, member(ident(BASH_ARGS), "length"), int(*n))
        }
        _ => ternary(
            binary(BinOp::Gt, offset.clone(), int(0)),
            binary(BinOp::Sub, offset.clone(), int(1)),
            binary(BinOp::Add, member(ident(BASH_ARGS), "length"), offset),
        ),
    }
}

fn positional_slice_expr(
    target: Expression,
    offset: Expression,
    length: Option<Expression>,
) -> Expression {
    match &offset.kind {
        ExprKind::Lit(Literal::Int(0)) => array_slice_expr(
            method(array(vec![positional("0")]), "concat", vec![target]),
            int(0),
            length,
        ),
        _ => array_slice_expr(target, positional_slice_start(offset), length),
    }
}

fn array_slice_expr(
    target: Expression,
    offset: Expression,
    length: Option<Expression>,
) -> Expression {
    let end = match length {
        None => member(target.clone(), "length"),
        Some(len) => match &len.kind {
            ExprKind::Lit(Literal::Int(n)) if *n < 0 => len,
            _ => binary(BinOp::Add, offset.clone(), len),
        },
    };
    call_named("__bash_slice", vec![target, offset, end])
}

fn array_map_expr(source: Expression, item: &str, _out: &str, mapped: Expression) -> Expression {
    method(
        source,
        "map",
        vec![lambda_expr(vec![param_named(item)], mapped)],
    )
}

/// Command substitution strips trailing newlines from the captured text.
fn strip_trailing_newlines(e: Expression) -> Expression {
    regex_replace_expr(e, "\\n+$", "", lit(""))
}

fn strip_command_substitution_value(e: Expression) -> Expression {
    let value = "__bash_capture";
    let (mut params, mut args) = shell_state_params_args();
    params.insert(0, param_named(value));
    args.insert(0, e);
    iife_with_args(
        vec![
            Statement::new(StmtKind::If {
                cond: regex_test_expr(regexp(lit("\\x00"), ""), ident(value)),
                then_body: vec![bash_stderr_stmt(lit(
                    "bash: warning: command substitution: ignored null byte in input\n",
                ))],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::Return(Some(strip_trailing_newlines(
                regex_replace_expr(ident(value), "\\x00", "g", lit("")),
            )))),
        ],
        params,
        args,
    )
}

fn command_substitution_expr(e: Expression) -> Expression {
    strip_command_substitution_value(e)
}

fn bash_stderr_write_expr(text: Expression) -> Expression {
    ternary(
        ident("__bash_stderr_null"),
        int(0),
        ternary(
            ident("__bash_stderr_to_stdout"),
            call_named("printf", vec![lit("%s"), text]),
            int(0),
        ),
    )
}

fn bash_stdout_write_expr(text: Expression) -> Expression {
    ternary(
        ident("__bash_stdout_to_stderr"),
        bash_stderr_write_expr(text.clone()),
        ternary(
            ident("__bash_stdout_null"),
            int(0),
            call_named("printf", vec![lit("%s"), text]),
        ),
    )
}

fn bash_stdout_status_expr(text: Expression) -> Expression {
    sequence(vec![bash_stdout_write_expr(text), int(0)])
}

fn bash_stderr_stmt(text: Expression) -> Statement {
    expr_stmt(bash_stderr_write_expr(text))
}

fn spawn_status(result: Expression) -> Expression {
    ternary(
        binary(BinOp::StrictNotEq, member(result.clone(), "error"), null()),
        int(127),
        binary(BinOp::NullCoalesce, member(result, "status"), int(127)),
    )
}
fn bash_spawn_status_expr(name: &str, args: Vec<Expression>) -> Expression {
    bash_spawn_status_argv_expr(name, array(args))
}

fn bash_spawn_status_argv_expr(name: &str, argv: Expression) -> Expression {
    let result = "__bash_spawn_result".to_string();
    iife(vec![
        let_stmt(
            &result,
            call_named(
                "__bash_spawn",
                vec![
                    lit(name),
                    argv,
                    object(vec![
                        ("input", ident("__bash_stdin")),
                        ("cwd", ident("PWD")),
                    ]),
                ],
            ),
        ),
        expr_stmt(bash_stdout_write_expr(member(ident(&result), "stdout"))),
        Statement::new(StmtKind::Return(Some(spawn_status(ident(&result))))),
    ])
}
fn bash_test_status_argv_expr(argv: Expression) -> Expression {
    bash_test_status_argv_expr_with_command(argv, "[")
}

fn bash_test_status_argv_expr_with_command(argv: Expression, command: &str) -> Expression {
    let args = "__bash_test_argv".to_string();
    let stderr_to_stdout = "__bash_test_stderr_to_stdout";
    let op = index(ident(&args), int(0));
    let second = index(ident(&args), int(1));
    let third = index(ident(&args), int(2));
    let is_binary_op = |expr: Expression| {
        binary(
            BinOp::Or,
            binary(BinOp::StrictEq, expr.clone(), lit("=")),
            binary(
                BinOp::Or,
                binary(BinOp::StrictEq, expr.clone(), lit("==")),
                binary(
                    BinOp::Or,
                    binary(BinOp::StrictEq, expr.clone(), lit("!=")),
                    binary(
                        BinOp::Or,
                        binary(BinOp::StrictEq, expr.clone(), lit("-eq")),
                        binary(
                            BinOp::Or,
                            binary(BinOp::StrictEq, expr.clone(), lit("-ne")),
                            binary(
                                BinOp::Or,
                                binary(BinOp::StrictEq, expr.clone(), lit("-lt")),
                                binary(
                                    BinOp::Or,
                                    binary(BinOp::StrictEq, expr.clone(), lit("-le")),
                                    binary(
                                        BinOp::Or,
                                        binary(BinOp::StrictEq, expr.clone(), lit("-gt")),
                                        binary(BinOp::StrictEq, expr, lit("-ge")),
                                    ),
                                ),
                            ),
                        ),
                    ),
                ),
            ),
        )
    };
    iife_with_args(
        vec![
            let_stmt(&args, argv),
            Statement::new(StmtKind::If {
                cond: binary(
                    BinOp::And,
                    binary(BinOp::StrictEq, member(ident(&args), "length"), int(2)),
                    binary(
                        BinOp::Or,
                        binary(BinOp::StrictEq, op.clone(), lit("=")),
                        binary(
                            BinOp::Or,
                            binary(BinOp::StrictEq, op.clone(), lit("==")),
                            binary(BinOp::StrictEq, op, lit("!=")),
                        ),
                    ),
                ),
                then_body: vec![
                    expr_stmt(ternary(
                        ident(stderr_to_stdout),
                        call_named(
                            "printf",
                            vec![lit("%s"), lit("bash: [: unary operator expected\n")],
                        ),
                        int(0),
                    )),
                    Statement::new(StmtKind::Return(Some(int(2)))),
                ],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::If {
                cond: binary(
                    BinOp::And,
                    binary(BinOp::Gt, member(ident(&args), "length"), int(3)),
                    binary(BinOp::Or, is_binary_op(second), is_binary_op(third)),
                ),
                then_body: vec![
                    expr_stmt(ternary(
                        ident(stderr_to_stdout),
                        call_named(
                            "printf",
                            vec![
                                lit("%s"),
                                lit(&format!("bash: {command}: too many arguments\n")),
                            ],
                        ),
                        int(0),
                    )),
                    Statement::new(StmtKind::Return(Some(int(2)))),
                ],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::Return(Some(bash_spawn_status_argv_expr(
                "test",
                ident(&args),
            )))),
        ],
        vec![param_named(stderr_to_stdout)],
        vec![ident("__bash_stderr_to_stdout")],
    )
}
fn bash_dynamic_command_status_expr(argv: Expression) -> Expression {
    let args = "__bash_dynamic_argv".to_string();
    let result = "__bash_dynamic_spawn".to_string();
    iife(vec![
        let_stmt(&args, argv),
        Statement::new(StmtKind::If {
            cond: binary(BinOp::StrictEq, member(ident(&args), "length"), int(0)),
            then_body: vec![Statement::new(StmtKind::Return(Some(int(0))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        let_stmt(
            &result,
            call_named(
                "__bash_spawn",
                vec![
                    method(ident(&args), "shift", vec![]),
                    ident(&args),
                    object(vec![
                        ("input", ident("__bash_stdin")),
                        ("cwd", ident("PWD")),
                    ]),
                ],
            ),
        ),
        Statement::new(StmtKind::Return(Some(spawn_status(ident(&result))))),
    ])
}
fn bash_spawn_test_expr(args: Vec<Expression>) -> Expression {
    binary(
        BinOp::StrictEq,
        bash_spawn_status_expr("test", args),
        int(0),
    )
}
fn shell_option_test_expr(option: &Expression, flags: &BTreeSet<char>) -> Expression {
    match literal_string(option) {
        Some("errexit") => Expression::bool(flags.contains(&'e')),
        Some("noglob") => Expression::bool(flags.contains(&'f')),
        Some("nounset") => Expression::bool(flags.contains(&'u')),
        Some("braceexpand") => Expression::bool(flags.contains(&'B')),
        Some("hashall") => Expression::bool(flags.contains(&'h')),
        Some("pipefail") => Expression::bool(flags.contains(&'P')),
        Some(_) => Expression::bool(false),
        None => Expression::bool(false),
    }
}
fn bash_env_value_expr(name: &str, fallback: Expression) -> Expression {
    let env = "__bash_env";
    let pair = "__bash_env_pair";
    iife(vec![
        let_stmt(env, call_named("__bash_getenv", vec![])),
        Statement::new(StmtKind::ForIn {
            var: pair.to_string(),
            key: None,
            iter: ident(env),
            body: vec![Statement::new(StmtKind::If {
                cond: binary(BinOp::StrictEq, index(ident(pair), int(0)), lit(name)),
                then_body: vec![Statement::new(StmtKind::Return(Some(index(
                    ident(pair),
                    int(1),
                ))))],
                elifs: Vec::new(),
                else_body: None,
            })],
            of: true,
            else_body: None,
            is_async: false,
        }),
        Statement::new(StmtKind::Return(Some(fallback))),
    ])
}
fn parameter_error_stmts(
    in_subshell: bool,
    (name, message): (String, Expression),
) -> Vec<Statement> {
    let mut stmts = vec![bash_stderr_stmt(call_named(
        "__bash_sprintf",
        vec![lit("%s: %s\n"), lit(&name), message],
    ))];
    stmts.push(if in_subshell {
        Statement::new(StmtKind::Return(Some(int(1))))
    } else {
        Statement::new(StmtKind::Exit {
            status: Some(int(1)),
        })
    });
    stmts
}
fn bad_substitution_expr() -> Expression {
    iife(vec![
        bash_stderr_stmt(lit("bash: bad substitution\n")),
        assign_stmt(ident("__bash_status"), int(1)),
        Statement::new(StmtKind::Return(Some(lit("")))),
    ])
}
fn param_value(e: Expression) -> Expression {
    if matches!(
        &e.kind,
        ExprKind::Lit(_) | ExprKind::Ident(_) | ExprKind::Index { .. } | ExprKind::Member { .. }
    ) {
        return binary(BinOp::NullCoalesce, e, lit(""));
    }
    iife(vec![
        let_stmt("__bash_param_value", e),
        Statement::new(StmtKind::Return(Some(binary(
            BinOp::NullCoalesce,
            ident("__bash_param_value"),
            lit(""),
        )))),
    ])
}
fn param_length(e: Expression) -> Expression {
    let value = param_value(e);
    ternary(
        binary(
            BinOp::StrictEq,
            Expression::new(ExprKind::TypeOf(Box::new(value.clone()))),
            lit("string"),
        ),
        call_named("__bash_strlen", vec![value.clone()]),
        call_named(
            "__bash_strlen",
            vec![call_named("__bash_string", vec![value])],
        ),
    )
}

fn param_count(e: Expression) -> Expression {
    let value = param_value(e);
    ternary(
        call_named("__bash_is_array", vec![value.clone()]),
        member(call_named("__bash_keys", vec![value.clone()]), "length"),
        member(call_named("__bash_dict_keys", vec![value]), "length"),
    )
}

fn literal_string(e: &Expression) -> Option<&str> {
    match &e.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s),
        _ => None,
    }
}

fn static_parts_text(parts: &[Part]) -> Option<String> {
    let mut out = String::new();
    for part in parts {
        match part {
            Part::Text(text) => out.push_str(text),
            Part::Expr(_) => return None,
        }
    }
    Some(out)
}

fn literal_key_string(e: &Expression) -> Option<String> {
    match &e.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.clone()),
        ExprKind::Lit(Literal::Int(n)) => Some(n.to_string()),
        ExprKind::Lit(Literal::Float(n)) if n.fract() == 0.0 => Some((*n as i64).to_string()),
        _ => None,
    }
}

fn bash_array_snapshot(e: &Expression, assoc: bool) -> Option<Vec<(String, String)>> {
    match &e.kind {
        ExprKind::Array(items) => items
            .iter()
            .enumerate()
            .map(|(idx, elem)| {
                if elem.spread {
                    return None;
                }
                literal_string(&elem.value).map(|value| (idx.to_string(), value.to_string()))
            })
            .collect(),
        ExprKind::Map(items) if assoc => items
            .iter()
            .map(|(key, value)| {
                Some((literal_key_string(key)?, literal_string(value)?.to_string()))
            })
            .collect(),
        ExprKind::Object(items) if assoc => items
            .iter()
            .map(|prop| match prop {
                ObjectProperty::KeyValue { key, value }
                | ObjectProperty::Computed { key, value } => {
                    Some((literal_key_string(key)?, literal_string(value)?.to_string()))
                }
                ObjectProperty::Shorthand(name) => Some((name.clone(), name.clone())),
                _ => None,
            })
            .collect(),
        _ => None,
    }
}

fn bash_array_declare_payload(entries: &[(String, String)], quote_values: bool) -> String {
    entries
        .iter()
        .map(|(key, value)| {
            let value = if quote_values {
                bash_double_quote(value)
            } else {
                value.clone()
            };
            format!("[{key}]={value}")
        })
        .collect::<Vec<_>>()
        .join(" ")
}

fn bash_double_quote(s: &str) -> String {
    format!(
        "\"{}\"",
        s.replace('\\', "\\\\")
            .replace('"', "\\\"")
            .replace('$', "\\$")
            .replace('`', "\\`")
    )
}

fn bash_single_quote(s: &str) -> String {
    if s.is_empty() {
        "''".into()
    } else {
        format!("'{}'", s.replace('\'', "'\\''"))
    }
}

fn bash_quote(s: &str) -> String {
    if s.is_empty() {
        return "''".into();
    }
    if s.chars().any(|c| c.is_control()) {
        return bash_ansi_c_quote(s);
    }
    bash_single_quote(s)
}

fn bash_ansi_c_quote(s: &str) -> String {
    let mut out = String::from("$'");
    for c in s.chars() {
        match c {
            '\n' => out.push_str("\\n"),
            '\t' => out.push_str("\\t"),
            '\r' => out.push_str("\\r"),
            '\u{7}' => out.push_str("\\a"),
            '\u{8}' => out.push_str("\\b"),
            '\u{c}' => out.push_str("\\f"),
            '\u{b}' => out.push_str("\\v"),
            '\\' => out.push_str("\\\\"),
            '\'' => out.push_str("\\'"),
            other if other.is_control() => out.push_str(&format!("\\x{:02x}", other as u32)),
            other => out.push(other),
        }
    }
    out.push('\'');
    out
}

fn flag_state(flags: &str, needle: char) -> Option<bool> {
    let mut sign = None;
    for c in flags.chars() {
        match c {
            '-' => sign = Some(true),
            '+' => sign = Some(false),
            _ if c == needle => return sign,
            _ => {}
        }
    }
    None
}

fn flag_enabled(flags: &str, needle: char) -> bool {
    flag_state(flags, needle) == Some(true)
}

fn flag_disabled(flags: &str, needle: char) -> bool {
    flag_state(flags, needle) == Some(false)
}

fn word_is_bash_command(pair: &Pair<Rule>) -> bool {
    let raw = pair.as_str().trim();
    matches!(
        raw,
        "bash" | "$BASH" | "${BASH}" | "\"$BASH\"" | "\"${BASH}\""
    )
}

/// A lowered command: its expression plus whether that expression is already
/// a boolean (tests, `(( ))`, `&&` chains) or an exit status (everything else).
#[derive(Clone)]
struct Cmd {
    expr: Expression,
    is_bool: bool,
}

#[derive(Clone)]
struct ShellArg {
    value: Expression,
    spread: bool,
}

impl Cmd {
    fn status(expr: Expression) -> Self {
        Cmd {
            expr,
            is_bool: false,
        }
    }
    fn boolean(expr: Expression) -> Self {
        Cmd {
            expr,
            is_bool: true,
        }
    }
    /// The command as a condition: success is true.
    fn cond(self) -> Expression {
        if self.is_bool {
            self.expr
        } else {
            unary(UnaryOp::Not, self.expr)
        }
    }
    /// The command as a statement that records `$?`.
    fn status_stmt(self) -> Statement {
        if self.is_bool {
            return assign_stmt(ident("__bash_status"), ternary(self.expr, int(0), int(1)));
        }
        match &self.expr.kind {
            ExprKind::Lit(Literal::Int(_)) => assign_stmt(ident("__bash_status"), self.expr),
            _ => expr_stmt(sequence(vec![
                assign_expr(ident("__bash_ret"), self.expr),
                assign_expr(ident("__bash_status"), status_of_value(ident("__bash_ret"))),
            ])),
        }
    }
}

/// What a simple command lowered to: plain statements, or one expression
/// whose value is the command's status/truth.
struct Lowered {
    stmts: Vec<Statement>,
    value: Option<Cmd>,
}
impl Lowered {
    fn stmts(stmts: Vec<Statement>) -> Self {
        Lowered { stmts, value: None }
    }
    fn value(cmd: Cmd) -> Self {
        Lowered {
            stmts: Vec::new(),
            value: Some(cmd),
        }
    }
    fn with_last_arg(mut self, arg: Expression) -> Self {
        let update = assign_expr(ident("__bash_last_arg"), arg.clone());
        if let Some(cmd) = &mut self.value {
            let expr = cmd.expr.clone();
            cmd.expr = sequence(vec![update, expr]);
        } else {
            self.stmts.push(assign_stmt(ident("__bash_last_arg"), arg));
        }
        self
    }
    fn into_stmts(self) -> Vec<Statement> {
        match self.value {
            Some(cmd) => vec![cmd.status_stmt()],
            None => self.stmts,
        }
    }
    fn into_expr(self) -> Cmd {
        match self.value {
            Some(cmd) => cmd,
            None => {
                let mut stmts = self.stmts;
                stmts.push(Statement::new(StmtKind::Return(Some(ident(
                    "__bash_status",
                )))));
                Cmd::status(shell_state_iife(stmts))
            }
        }
    }
    fn into_current_expr(self) -> Cmd {
        match self.value {
            Some(cmd) => cmd,
            None => {
                let mut stmts = self.stmts;
                if let Some(mut exprs) = stmts_to_exprs(&stmts) {
                    exprs.push(ident("__bash_status"));
                    return Cmd::status(sequence(exprs));
                }
                stmts.push(Statement::new(StmtKind::Return(Some(ident(
                    "__bash_status",
                )))));
                Cmd::status(iife(stmts))
            }
        }
    }
}

/// One stage of a pipeline: a command with the redirections written after it.
struct Unit<'i> {
    cmd: Pair<'i, Rule>,
    redirs: Vec<Pair<'i, Rule>>,
    pipe_stderr: bool,
}

// ── walker ───────────────────────────────────────────────────────────────────

#[derive(Clone)]
struct Walker {
    heredocs: VecDeque<HereDoc>,
    counter: usize,
    functions: HashMap<String, String>,
    function_bodies: HashMap<String, String>,
    function_commands: HashMap<String, String>,
    variables: BTreeSet<String>,
    variable_values: HashMap<String, String>,
    array_values: HashMap<String, Vec<(String, String)>>,
    indexed_arrays: HashSet<String>,
    assoc_arrays: HashSet<String>,
    readonly_vars: HashSet<String>,
    integer_vars: HashSet<String>,
    lowercase_vars: HashSet<String>,
    uppercase_vars: HashSet<String>,
    exported_vars: HashSet<String>,
    exported_functions: HashSet<String>,
    traced_functions: HashSet<String>,
    namerefs: HashMap<String, String>,
    aliases: HashMap<String, String>,
    positional_values: Vec<String>,
    brace_expansion_enabled: bool,
    nullglob_enabled: bool,
    dotglob_enabled: bool,
    failglob_enabled: bool,
    globstar_enabled: bool,
    expand_aliases_enabled: bool,
    nocaseglob_enabled: bool,
    nocasematch_enabled: bool,
    patsub_replacement_enabled: bool,
    lastpipe_enabled: bool,
    shell_flags: BTreeSet<char>,
    subshell_depth: usize,
    arith_stack: Vec<String>,
    dynamic_arith: bool,
    function_depth: usize,
    inline_call_context: bool,
    inline_stack: Vec<String>,
    suppress_function_inlining: bool,
    hoisted_functions: Vec<Statement>,
    exit_trap_body: Option<Vec<Statement>>,
}

type R<T> = Result<T, String>;

impl Walker {
    fn fresh(&mut self, base: &str) -> String {
        self.counter += 1;
        format!("{base}_{}", self.counter)
    }

    fn record_variable(&mut self, name: &str) {
        if is_name(name) {
            self.variables.insert(name.to_string());
        }
    }

    fn record_array_element_value(&mut self, name: &str, key: String, value: String) {
        let entries = self.array_values.entry(name.to_string()).or_default();
        if let Some((_, existing)) = entries.iter_mut().find(|(k, _)| k == &key) {
            *existing = value;
        } else {
            entries.push((key, value));
        }
    }

    fn static_array_keys(&self, name: &str) -> Option<Vec<String>> {
        let mut keys: Vec<String> = self
            .array_values
            .get(name)?
            .iter()
            .map(|(key, _)| key.clone())
            .collect();
        if self.indexed_arrays.contains(name) {
            keys.sort_by(|a, b| {
                let an = a.parse::<i64>().ok();
                let bn = b.parse::<i64>().ok();
                match (an, bn) {
                    (Some(a), Some(b)) => a.cmp(&b),
                    _ => a.cmp(b),
                }
            });
        } else {
            keys.sort();
        }
        keys.dedup();
        Some(keys)
    }

    fn nameref_would_cycle(&self, name: &str, target: &str) -> bool {
        if !is_name(target) {
            return false;
        }
        let mut current = target.to_string();
        let mut seen = HashSet::new();
        while seen.insert(current.clone()) {
            if current == name {
                return true;
            }
            let Some(next) = self.namerefs.get(&current) else {
                return false;
            };
            if !is_name(next) {
                return false;
            }
            current = next.clone();
        }
        false
    }

    fn static_ifs_first_char(&self) -> Option<String> {
        self.variable_values
            .get("IFS")
            .and_then(|s| s.chars().next())
            .map(|c| c.to_string())
    }

    fn variable_is_set(&self, name: &str) -> bool {
        let resolved = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        self.variables.contains(&resolved)
            || self.variable_values.contains_key(&resolved)
            || self.array_values.contains_key(&resolved)
            || self.indexed_arrays.contains(&resolved)
            || self.assoc_arrays.contains(&resolved)
    }

    fn variable_exists_expr(&mut self, name: &str) -> R<Expression> {
        if let Ok(pos) = name.parse::<usize>() {
            if pos == 0 {
                return Ok(Expression::bool(true));
            }
            return Ok(binary(
                BinOp::Lt,
                int(pos as i64 - 1),
                member(ident(BASH_ARGS), "length"),
            ));
        }
        let resolved = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        if let Some(base) = resolved
            .strip_suffix("[@]")
            .or_else(|| resolved.strip_suffix("[*]"))
        {
            return Ok(binary(BinOp::Gt, member(ident(base), "length"), int(0)));
        }
        if !resolved.contains('[')
            && (self.array_values.contains_key(&resolved)
                || self.indexed_arrays.contains(&resolved)
                || self.assoc_arrays.contains(&resolved))
        {
            return Ok(binary(
                BinOp::StrictNotEq,
                index(ident(&resolved), int(0)),
                undefined(),
            ));
        }
        if is_unset_target(&resolved) {
            return Ok(binary(
                BinOp::StrictNotEq,
                self.name_or_element_target(&resolved)?,
                undefined(),
            ));
        }
        Ok(Expression::bool(false))
    }

    fn shell_flags_text(&self) -> String {
        self.shell_flags.iter().collect()
    }

    fn target_text_is_readonly(&self, text: &str) -> bool {
        let base = text.split_once('[').map(|(name, _)| name).unwrap_or(text);
        if is_readonly_parameter_target(base) {
            return true;
        }
        let resolved = self
            .resolve_nameref_text(base)
            .unwrap_or_else(|| base.to_string());
        self.readonly_vars.contains(&resolved)
    }

    fn assignment_word_targets_readonly(&self, pair: &Pair<Rule>) -> bool {
        assignment_word_target_text(pair)
            .is_some_and(|target| self.target_text_is_readonly(&target))
    }

    fn assignment_word_has_invalid_indirect(&self, pair: &Pair<Rule>) -> bool {
        let mut rest = pair.as_str();
        while let Some(start) = rest.find("${!") {
            rest = &rest[start + 3..];
            let Some(end) = rest.find('}') else {
                return true;
            };
            let base = &rest[..end];
            rest = &rest[end + 1..];
            let Some(target_name) = base
                .split([':', '-', '+', '=', '?', '/', '#', '%', '@'])
                .next()
                .filter(|name| is_name(name))
            else {
                continue;
            };
            let Some(target) = self.variable_values.get(target_name) else {
                continue;
            };
            if target.is_empty() || !is_valid_indirect_target(target) {
                return true;
            }
        }
        false
    }

    fn apply_variable_attributes(&mut self, name: &str, value: Expression) -> R<Expression> {
        let resolved = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        Ok(if self.integer_vars.contains(&resolved) {
            let numeric = if let Some(s) = literal_string(&value) {
                self.parse_arith_text(s)
                    .unwrap_or_else(|_| to_number(value))
            } else {
                to_number(value)
            };
            call_named("__bash_string", vec![numeric])
        } else if self.lowercase_vars.contains(&resolved) {
            method(param_value(value), "toLowerCase", vec![])
        } else if self.uppercase_vars.contains(&resolved) {
            method(param_value(value), "toUpperCase", vec![])
        } else {
            value
        })
    }

    fn word_has_nounset_reference(&self, pair: &Pair<Rule>) -> bool {
        match pair.as_rule() {
            Rule::simple_param => {
                let mut inner = pair.clone().into_inner();
                let Some(p) = inner.next() else {
                    return false;
                };
                p.as_rule() == Rule::name && self.name_is_unset_for_nounset(p.as_str())
            }
            Rule::braced_param => simple_braced_param_name(pair.as_str())
                .is_some_and(|name| self.name_is_unset_for_nounset(name)),
            _ => pair
                .clone()
                .into_inner()
                .any(|p| self.word_has_nounset_reference(&p)),
        }
    }

    fn name_is_unset_for_nounset(&self, name: &str) -> bool {
        if !is_name(name) {
            return false;
        }
        if matches!(
            name,
            "BASH" | "HOME" | "PWD" | "OLDPWD" | "BASH_EXECUTION_STRING"
        ) {
            return false;
        }
        let resolved = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        !self.variables.contains(&resolved) && !self.variable_values.contains_key(&resolved)
    }

    // ── lists and pipelines ──────────────────────────────────────────────

    fn walk_list(&mut self, pair: Pair<Rule>) -> R<Vec<Statement>> {
        let mut out = Vec::new();
        let parts: Vec<Pair<Rule>> = pair.into_inner().collect();
        let mut current_line: Option<usize> = None;
        for i in 0..parts.len() {
            if parts[i].as_rule() == Rule::and_or {
                let line = parts[i].as_span().start_pos().line_col().0;
                if current_line != Some(line) {
                    out.push(assign_stmt(ident("__bash_abort_line"), Expression::bool(false)));
                    current_line = Some(line);
                }
                let background = parts
                    .get(i + 1)
                    .is_some_and(|p| p.as_rule() == Rule::background_op);
                let mut body;
                if background {
                    let mut child = self.clone();
                    child.subshell_depth += 1;
                    body = child.walk_and_or(parts[i].clone())?;
                    self.heredocs = child.heredocs.clone();
                    self.counter = child.counter;
                    rewrite_exit_to_return(&mut body);
                    body.push(Statement::new(StmtKind::Return(Some(ident(
                        "__bash_status",
                    )))));
                    let (params, args) = self.shell_child_bindings(&child, true, &[], None);
                    body = vec![assign_stmt(
                        ident("__bash_status"),
                        iife_with_args(body, params, args),
                    )];
                    body.push(assign_stmt(
                        ident("__bash_job_seq"),
                        binary(BinOp::Add, ident("__bash_job_seq"), int(1)),
                    ));
                    body.push(assign_stmt(
                        ident("__bash_last_bg_pid"),
                        ident("__bash_job_seq"),
                    ));
                    body.push(assign_stmt(
                        index(ident("__bash_jobs"), ident("__bash_last_bg_pid")),
                        ident("__bash_status"),
                    ));
                    body.push(assign_stmt(ident("__bash_status"), int(0)));
                } else {
                    body = self.walk_and_or(parts[i].clone())?;
                }
                out.push(Statement::new(StmtKind::If {
                    cond: unary(UnaryOp::Not, ident("__bash_abort_line")),
                    then_body: body,
                    elifs: Vec::new(),
                    else_body: None,
                }));
            }
        }
        Ok(out)
    }

    /// A list as one expression, for conditions: `if a; b; then` tests `b`.
    fn list_expr(&mut self, pair: Pair<Rule>) -> R<Cmd> {
        let mut cmds: Vec<Cmd> = Vec::new();
        for inner in pair.into_inner() {
            if inner.as_rule() == Rule::and_or {
                cmds.push(self.and_or_expr(inner)?);
            }
        }
        let last = cmds.pop().ok_or("empty condition list")?;
        if cmds.is_empty() {
            return Ok(last);
        }
        let mut seq: Vec<Expression> = cmds.into_iter().map(|c| c.expr).collect();
        let is_bool = last.is_bool;
        seq.push(last.expr);
        Ok(Cmd {
            expr: sequence(seq),
            is_bool,
        })
    }

    fn walk_and_or(&mut self, pair: Pair<Rule>) -> R<Vec<Statement>> {
        let parts: Vec<Pair<Rule>> = pair.clone().into_inner().collect();
        let has_ops = parts
            .iter()
            .any(|p| matches!(p.as_rule(), Rule::and_op | Rule::or_op));
        if !has_ops {
            let mut out = Vec::new();
            for p in parts {
                if p.as_rule() == Rule::pipeline {
                    out.extend(self.walk_pipeline_stmts(p)?);
                }
            }
            return Ok(out);
        }
        self.and_or_stmts(pair)
    }

    /// `a && b || c` as shell command sequencing, left-associative.
    fn and_or_stmts(&mut self, pair: Pair<Rule>) -> R<Vec<Statement>> {
        let mut out = Vec::new();
        let mut pending_op: Option<BinOp> = None;
        let mut saw_pipeline = false;
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::and_op => pending_op = Some(BinOp::And),
                Rule::or_op => pending_op = Some(BinOp::Or),
                Rule::pipeline => {
                    let body = self.walk_pipeline_stmts(inner)?;
                    if !saw_pipeline {
                        out.extend(body);
                        saw_pipeline = true;
                    } else {
                        let op = pending_op.take().ok_or("missing &&/|| operator")?;
                        let cond = match op {
                            BinOp::And => binary(BinOp::StrictEq, ident("__bash_status"), int(0)),
                            BinOp::Or => binary(BinOp::NotEq, ident("__bash_status"), int(0)),
                            _ => unreachable!(),
                        };
                        out.push(Statement::new(StmtKind::If {
                            cond,
                            then_body: body,
                            elifs: Vec::new(),
                            else_body: None,
                        }));
                    }
                }
                _ => {}
            }
        }
        if saw_pipeline {
            Ok(out)
        } else {
            Err("empty and_or".into())
        }
    }

    /// `a && b || c` as an expression, for condition contexts.
    fn and_or_expr(&mut self, pair: Pair<Rule>) -> R<Cmd> {
        let parts: Vec<Pair<Rule>> = pair.clone().into_inner().collect();
        if parts
            .iter()
            .any(|p| matches!(p.as_rule(), Rule::and_op | Rule::or_op))
        {
            let mut body = self.and_or_stmts(pair)?;
            body.push(Statement::new(StmtKind::Return(Some(ident(
                "__bash_status",
            )))));
            return Ok(Cmd::status(iife(body)));
        }
        let mut result: Option<Cmd> = None;
        let mut pending_op: Option<BinOp> = None;
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::and_op => pending_op = Some(BinOp::And),
                Rule::or_op => pending_op = Some(BinOp::Or),
                Rule::pipeline => {
                    let cmd = self.pipeline_expr(inner)?;
                    result = Some(match (result, pending_op.take()) {
                        (None, _) => cmd,
                        (Some(left), Some(op)) => Cmd::boolean(binary(op, left.cond(), cmd.cond())),
                        (Some(_), None) => cmd,
                    });
                }
                _ => {}
            }
        }
        result.ok_or_else(|| "empty and_or".to_string())
    }

    /// Split a pipeline into `!` negation and its stages with their redirections.
    fn pipeline_units<'i>(&mut self, pair: Pair<'i, Rule>) -> (bool, Option<bool>, Vec<Unit<'i>>) {
        let mut negate = false;
        let mut time_posix = None;
        let mut units: Vec<Unit<'i>> = Vec::new();
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::bang => negate = !negate,
                Rule::time_prefix => time_posix = Some(inner.as_str().contains("-p")),
                Rule::pipe_op => {
                    if let Some(u) = units.last_mut() {
                        u.pipe_stderr = inner.as_str() == "|&";
                    }
                }
                Rule::redirection => {
                    if let Some(u) = units.last_mut() {
                        u.redirs.push(inner);
                    }
                }
                _ => units.push(Unit {
                    cmd: inner,
                    redirs: Vec::new(),
                    pipe_stderr: false,
                }),
            }
        }
        (negate, time_posix, units)
    }

    fn walk_pipeline_stmts(&mut self, pair: Pair<Rule>) -> R<Vec<Statement>> {
        let (negate, time_posix, mut units) = self.pipeline_units(pair);
        if units.len() == 1 && !negate && time_posix.is_none() {
            let unit = units.remove(0);
            return Ok(self.unit_lowered(unit)?.into_stmts());
        }
        let cmd = self.pipeline_from_units(negate, time_posix, units)?;
        Ok(vec![cmd.status_stmt()])
    }

    fn pipeline_expr(&mut self, pair: Pair<Rule>) -> R<Cmd> {
        let (negate, time_posix, units) = self.pipeline_units(pair);
        self.pipeline_from_units(negate, time_posix, units)
    }

    fn pipeline_from_units(
        &mut self,
        negate: bool,
        time_posix: Option<bool>,
        mut units: Vec<Unit>,
    ) -> R<Cmd> {
        let mut cmd = if units.len() == 1 {
            self.unit_lowered(units.remove(0))?.into_current_expr()
        } else {
            // Each stage's output becomes the next stage's stdin; the last
            // stage's status is the pipeline's.
            let saved = self.fresh("__bash_saved");
            let mut stmts = vec![let_stmt(&saved, ident("__bash_stdin"))];
            stmts.push(expr_stmt(method(
                ident("PIPESTATUS"),
                "splice",
                vec![int(0), member(ident("PIPESTATUS"), "length")],
            )));
            let n = units.len();
            for (i, unit) in units.into_iter().enumerate() {
                let pipe_stderr = unit.pipe_stderr;
                if i + 1 < n {
                    let (stage, params, args) = self.isolated_pipeline_stage(unit)?;
                    let captured = if pipe_stderr {
                        let saved_stderr = self.fresh("__bash_saved_stderr_to_stdout");
                        let mut body = vec![
                            let_stmt(&saved_stderr, ident("__bash_stderr_to_stdout")),
                            assign_stmt(ident("__bash_stderr_to_stdout"), Expression::bool(true)),
                        ];
                        body.extend(stage);
                        body.push(assign_stmt(
                            ident("__bash_stderr_to_stdout"),
                            ident(&saved_stderr),
                        ));
                        capture_with_args(body, params, args)
                    } else {
                        capture_with_args(stage, params, args)
                    };
                    stmts.push(assign_stmt(ident("__bash_stdin"), captured));
                    stmts.push(assign_stmt(
                        index(ident("PIPESTATUS"), int(i as i64)),
                        ident("__bash_status"),
                    ));
                } else {
                    if self.lastpipe_enabled {
                        stmts.extend(self.unit_lowered(unit)?.into_stmts());
                    } else {
                        let (stage, params, args) = self.isolated_pipeline_stage(unit)?;
                        stmts.push(assign_stmt(
                            ident("__bash_status"),
                            iife_with_args(stage, params, args),
                        ));
                    }
                    stmts.push(assign_stmt(
                        index(ident("PIPESTATUS"), int(i as i64)),
                        ident("__bash_status"),
                    ));
                }
            }
            stmts.push(assign_stmt(ident("__bash_stdin"), ident(&saved)));
            let expr = if let Some(mut exprs) = stmts_to_exprs(&stmts) {
                exprs.push(ident("__bash_status"));
                sequence(exprs)
            } else {
                stmts.push(Statement::new(StmtKind::Return(Some(ident(
                    "__bash_status",
                )))));
                shell_state_iife(stmts)
            };
            Cmd::status(expr)
        };
        if negate {
            cmd = if cmd.is_bool {
                Cmd::boolean(unary(UnaryOp::Not, cmd.expr))
            } else {
                Cmd::boolean(binary(BinOp::NotEq, cmd.expr, int(0)))
            };
        }
        let _ = time_posix;
        Ok(cmd)
    }

    fn isolated_pipeline_stage(
        &mut self,
        unit: Unit,
    ) -> R<(Vec<Statement>, Vec<Param>, Vec<Expression>)> {
        let mut child = self.clone();
        child.subshell_depth += 1;
        let mut body = child.unit_lowered(unit)?.into_stmts();
        self.heredocs = child.heredocs.clone();
        self.counter = child.counter;
        rewrite_exit_to_return(&mut body);
        body.push(Statement::new(StmtKind::Return(Some(ident(
            "__bash_status",
        )))));
        let (params, args) = self.shell_child_bindings(&child, true, &[], None);
        Ok((body, params, args))
    }

    // ── commands ─────────────────────────────────────────────────────────

    /// A pipeline stage with its trailing redirections applied.
    fn unit_lowered(&mut self, unit: Unit) -> R<Lowered> {
        let Unit { cmd, redirs, .. } = unit;
        if cmd.as_rule() == Rule::simple_command {
            return self.walk_simple_command(cmd, redirs);
        }
        if cmd.as_rule() == Rule::while_clause && !redirs.is_empty() {
            if let Some(lowered) = self.walk_while_read_with_redirection(cmd.clone(), &redirs)? {
                return Ok(lowered);
            }
        }
        let lowered = self.compound_lowered(cmd)?;
        if redirs.is_empty() {
            return Ok(lowered);
        }
        let stmts = lowered.into_stmts();
        Ok(Lowered::stmts(self.apply_redirections(stmts, &redirs)?))
    }

    fn compound_lowered(&mut self, pair: Pair<Rule>) -> R<Lowered> {
        Ok(match pair.as_rule() {
            Rule::function_def => Lowered::stmts(vec![self.walk_function_def(pair)?]),
            Rule::brace_group => Lowered::stmts(vec![Statement::new(StmtKind::Block(
                self.walk_group(pair)?,
            ))]),
            Rule::subshell => {
                let mut child = self.clone();
                child.subshell_depth += 1;
                let mut body = child.walk_group(pair)?;
                self.heredocs = child.heredocs.clone();
                self.counter = child.counter;
                rewrite_exit_to_return(&mut body);
                body = self.catch_subshell_exit_throw(body);
                if let Some(trap_body) = child.exit_trap_body.clone() {
                    body.extend(trap_body);
                }
                body.push(Statement::new(StmtKind::Return(Some(ident(
                    "__bash_status",
                )))));
                let (params, args) = self.shell_child_bindings(&child, true, &[], None);
                Lowered::value(Cmd::status(iife_with_args(body, params, args)))
            }
            Rule::if_clause => Lowered::stmts(vec![self.walk_if(pair)?]),
            Rule::while_clause => Lowered::stmts(vec![self.walk_while(pair, false)?]),
            Rule::until_clause => Lowered::stmts(vec![self.walk_while(pair, true)?]),
            Rule::for_clause => Lowered::stmts(vec![self.walk_for_in(pair)?]),
            Rule::for_arithmetic => Lowered::stmts(vec![self.walk_for_arith(pair)?]),
            Rule::case_clause => Lowered::stmts(self.walk_case(pair)?),
            Rule::select_clause => Lowered::stmts(vec![self.walk_select(pair)?]),
            Rule::arithmetic_cmd => {
                if let Some(inner) = arithmetic_cmd_inner(pair.as_str())
                    && let Some(message) = bash_arithmetic_error_message(inner)
                {
                    Lowered::value(Cmd::status(sequence(vec![
                        bash_stderr_write_expr(lit(&format!("bash: {message}\n"))),
                        int(1),
                    ])))
                } else if let Some(inner) = arithmetic_cmd_inner(pair.as_str())
                    && arith_assignment_lhs(inner)
                        .as_deref()
                        .is_some_and(|name| self.target_text_is_readonly(name))
                {
                    Lowered::value(Cmd::status(int(1)))
                } else {
                    let v = self.arith_cmd_value(pair)?;
                    Lowered::value(Cmd::boolean(arith_bool(v)))
                }
            }
            Rule::arithmetic_cmd_error => {
                if let Some(expr) = self.simple_arithmetic_update_expr(pair.as_str()) {
                    Lowered::value(Cmd::boolean(arith_bool(expr)))
                } else {
                    let message = arithmetic_cmd_inner(pair.as_str())
                        .and_then(bash_arithmetic_error_message)
                        .unwrap_or("operand expected");
                    Lowered::value(Cmd::status(sequence(vec![
                        bash_stderr_write_expr(lit(&format!("bash: {message}\n"))),
                        int(1),
                    ])))
                }
            }
            Rule::test_double_bracket_unary_error => Lowered::value(Cmd::status(sequence(vec![
                bash_stderr_write_expr(lit(
                    "bash: unexpected argument `]]' to conditional unary operator\n",
                )),
                int(2),
            ]))),
            Rule::test_single_bracket | Rule::test_double_bracket => {
                if pair.as_rule() == Rule::test_single_bracket {
                    if let Some(expr) = self.single_bracket_one_arg_expr(&pair)? {
                        Lowered::value(Cmd::boolean(expr))
                    } else if let Some(expr) = single_bracket_raw_arith_error_status_expr(pair.as_str(), "[")
                        .or_else(|| single_bracket_pair_arith_error_status_expr(&pair, "["))
                    {
                        Lowered::value(Cmd::status(expr))
                    } else if let Some(argv) = self.single_bracket_runtime_argv(&pair)? {
                        Lowered::value(Cmd::status(bash_test_status_argv_expr(argv)))
                    } else if let Some(args) = single_bracket_literal_arith_args(&pair) {
                        Lowered::value(Cmd::status(bash_spawn_status_expr("test", args)))
                    } else {
                        Lowered::value(Cmd::boolean(self.walk_test_bracket(pair)?))
                    }
                } else {
                    if double_bracket_has_unary_without_operand(&pair) {
                        Lowered::value(Cmd::status(sequence(vec![
                            bash_stderr_write_expr(lit(
                                "bash: unexpected argument `]]' to conditional unary operator\n",
                            )),
                            int(2),
                        ])))
                    } else if double_bracket_uses_legacy_boolean_op(&pair) {
                        Lowered::value(Cmd::status(sequence(vec![
                            bash_stderr_write_expr(lit(
                                "bash: syntax error in conditional expression\n",
                            )),
                            int(2),
                        ])))
                    } else {
                        Lowered::value(Cmd::boolean(self.walk_test_bracket(pair)?))
                    }
                }
            }
            Rule::coproc_command => {
                let mut name = "COPROC".to_string();
                let mut body_pair = None;
                for inner in pair.clone().into_inner() {
                    match inner.as_rule() {
                        Rule::kw_coproc => {}
                        Rule::name => name = inner.as_str().to_string(),
                        _ => body_pair = Some(inner),
                    }
                }
                let deferred = body_pair
                    .as_ref()
                    .is_some_and(|p| source_has_command_word(p.as_str(), "read"));
                let mut child = self.clone();
                child.subshell_depth += 1;
                let body = match body_pair {
                    Some(body_pair) => child.command_stmts(body_pair)?,
                    None => Vec::new(),
                };
                self.heredocs = child.heredocs.clone();
                self.counter = child.counter;
                let (params, args) = self.shell_child_bindings(&child, true, &[], None);
                let read_fd = (64 + self.counter).to_string();
                self.counter += 1;
                let write_fd = (64 + self.counter).to_string();
                self.counter += 1;
                let mut stmts = vec![
                    assign_stmt(
                        ident("__bash_job_seq"),
                        binary(BinOp::Add, ident("__bash_job_seq"), int(1)),
                    ),
                    assign_stmt(ident("__bash_last_bg_pid"), ident("__bash_job_seq")),
                    assign_stmt(ident(&format!("{name}_PID")), ident("__bash_last_bg_pid")),
                    assign_stmt(ident(&name), array(vec![lit(&read_fd), lit(&write_fd)])),
                    assign_stmt(
                        index(ident("__bash_coprocs"), lit(&write_fd)),
                        object(vec![
                            ("read_fd", lit(&read_fd)),
                            ("pid", ident("__bash_last_bg_pid")),
                            ("body", lambda_block(body.clone())),
                        ]),
                    ),
                ];
                if deferred {
                    stmts.push(assign_stmt(
                        index(ident("__bash_fds"), lit(&read_fd)),
                        lit(""),
                    ));
                    stmts.push(assign_stmt(
                        index(ident("__bash_jobs"), ident("__bash_last_bg_pid")),
                        int(0),
                    ));
                    stmts.push(assign_stmt(ident("__bash_status"), int(0)));
                } else {
                    stmts.push(assign_stmt(
                        index(ident("__bash_fds"), lit(&read_fd)),
                        capture_with_args(body, params, args),
                    ));
                    stmts.push(assign_stmt(
                        index(ident("__bash_jobs"), ident("__bash_last_bg_pid")),
                        ident("__bash_status"),
                    ));
                }
                Lowered::stmts(stmts)
            }
            Rule::redirection => Lowered::stmts(Vec::new()),
            other => {
                return Err(format!(
                    "unsupported command rule {other:?}: {}",
                    pair.as_str()
                ));
            }
        })
    }

    /// Any command pair as statements (no outer redirections).
    fn command_stmts(&mut self, pair: Pair<Rule>) -> R<Vec<Statement>> {
        if pair.as_rule() == Rule::simple_command {
            return Ok(self.walk_simple_command(pair, Vec::new())?.into_stmts());
        }
        Ok(self.compound_lowered(pair)?.into_stmts())
    }

    fn walk_group(&mut self, pair: Pair<Rule>) -> R<Vec<Statement>> {
        let mut out = Vec::new();
        for inner in pair.into_inner() {
            if inner.as_rule() == Rule::list {
                out.extend(self.walk_list(inner)?);
            }
        }
        Ok(out)
    }

    // ── simple commands ──────────────────────────────────────────────────

    fn walk_simple_command<'i>(
        &mut self,
        pair: Pair<'i, Rule>,
        mut redirs: Vec<Pair<'i, Rule>>,
    ) -> R<Lowered> {
        let mut prefixes: Vec<Pair<Rule>> = Vec::new();
        let mut modifiers: Vec<String> = Vec::new();
        let mut cmd_word: Option<Pair<Rule>> = None;
        let mut suffix: Vec<Pair<Rule>> = Vec::new();
        let mut inner_redirs: Vec<Pair<Rule>> = Vec::new();
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::cmd_prefix => {
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::assignment_word {
                            prefixes.push(p);
                        }
                    }
                }
                Rule::cmd_modifier => {
                    if let Some(name) = command_modifier_name(&inner) {
                        modifiers.push(name);
                    }
                }
                Rule::cmd_word => cmd_word = inner.into_inner().next(),
                Rule::cmd_suffix => {
                    for item in inner.into_inner() {
                        match item.as_rule() {
                            Rule::redirection => inner_redirs.push(item),
                            _ => suffix.push(item),
                        }
                    }
                }
                Rule::redirection => inner_redirs.push(inner),
                _ => {}
            }
        }
        // Redirections written inside the command come before ones after it.
        inner_redirs.append(&mut redirs);
        let redirs = inner_redirs;

        let mut out = Vec::new();
        let Some(word) = cmd_word else {
            if modifiers.iter().any(|m| m == "exec") && !redirs.is_empty() {
                let mut stmts = out;
                stmts.extend(self.persistent_redirections(&redirs)?);
                return Ok(Lowered::stmts(stmts));
            }
            // Assignment-only command: `x=1 y+=2 arr=(a b)`.
            for p in prefixes {
                if let Some(error) = self.triggered_parameter_error(&p)? {
                    let stmts = parameter_error_stmts(self.subshell_depth > 0, error);
                    return Ok(Lowered::stmts(if redirs.is_empty() {
                        stmts
                    } else {
                        self.apply_redirections(stmts, &redirs)?
                    }));
                }
                if self.assignment_word_targets_readonly(&p) {
                    out.push(assign_stmt(ident("__bash_status"), int(1)));
                    continue;
                }
                let invalid_indirect = self.assignment_word_has_invalid_indirect(&p);
                let has_command_substitution = assignment_word_has_command_substitution(&p);
                out.push(self.walk_assignment(p)?);
                if invalid_indirect {
                    out.push(assign_stmt(ident("__bash_status"), int(1)));
                } else if !has_command_substitution {
                    out.push(assign_stmt(ident("__bash_status"), int(0)));
                }
            }
            if out.is_empty() && !redirs.is_empty() {
                // `> file` alone truncates the file.
                return Ok(Lowered::stmts(
                    self.apply_redirections(Vec::new(), &redirs)?,
                ));
            }
            return Ok(Lowered::stmts(out));
        };

        // `VAR=v cmd` scopes the assignment to the command.
        let mut ifs_empty = false;
        let mut scoped_prefixes: Vec<(String, Statement)> = Vec::new();
        let mut scoped_prefix_values: Vec<(String, Expression)> = Vec::new();
        for p in prefixes {
            if p.as_str().starts_with("IFS=") && p.as_str().trim() == "IFS=" {
                ifs_empty = true;
                continue;
            }
            let (target, op, value) = self.assignment_parts(p)?;
            match target {
                AssignTarget::Name(n) => {
                    if self.target_text_is_readonly(&n) {
                        out.push(assign_stmt(ident("__bash_status"), int(1)));
                        continue;
                    }
                    self.record_variable(&n);
                    let prefix_value = value.clone();
                    let stmt = if op == "+=" {
                        compound_append(self.name_or_element_target(&n)?, value)
                    } else {
                        assign_stmt(self.name_or_element_target(&n)?, value)
                    };
                    scoped_prefix_values.push((n.clone(), prefix_value));
                    scoped_prefixes.push((n, stmt));
                }
                AssignTarget::Element(arr, key) => {
                    if self.target_text_is_readonly(&arr) {
                        out.push(assign_stmt(ident("__bash_status"), int(1)));
                    } else {
                        out.push(assign_stmt(index(ident(&arr), key), value));
                    }
                }
            }
        }

        if self.shell_flags.contains(&'u')
            && (self.word_has_nounset_reference(&word)
                || suffix.iter().any(|w| self.word_has_nounset_reference(w)))
        {
            return Ok(Lowered::stmts(vec![Statement::new(StmtKind::Exit {
                status: Some(int(1)),
            })]));
        }

        let failglob_words = std::iter::once(word.clone())
            .chain(suffix.iter().cloned())
            .collect::<Vec<_>>();
        let name = literal_word_text(&word);
        let suppress_alias = word_starts_with_backslash(&word);
        if modifiers.iter().any(|m| m == "exec")
            && !redirs.is_empty()
            && suffix.is_empty()
            && redirs.iter().all(redirection_is_stdin_form)
        {
            let fd = word_source_text(&word);
            if !fd.is_empty() && fd.chars().all(|c| c.is_ascii_digit()) {
                let mut stmts = out;
                let body = self.persistent_redirections_with_fd_override(&redirs, Some(fd))?;
                if scoped_prefixes.is_empty() {
                    stmts.extend(body);
                } else {
                    stmts.extend(self.scoped_prefix_stmts(scoped_prefixes, body));
                }
                return Ok(Lowered::stmts(stmts));
            }
        }
        let exec_fd_override = if name.as_deref() == Some("exec")
            && !redirs.is_empty()
            && suffix.len() == 1
            && redirs.iter().all(redirection_is_stdin_form)
        {
            let fd = word_source_text(&suffix[0]);
            (!fd.is_empty() && fd.chars().all(|c| c.is_ascii_digit())).then_some(fd)
        } else {
            None
        };
        if name.as_deref() == Some("exec")
            && !redirs.is_empty()
            && (suffix.is_empty() || exec_fd_override.is_some())
        {
            let mut stmts = out;
            let body = self.persistent_redirections_with_fd_override(&redirs, exec_fd_override)?;
            if scoped_prefixes.is_empty() {
                stmts.extend(body);
            } else {
                stmts.extend(self.scoped_prefix_stmts(scoped_prefixes, body));
            }
            return Ok(Lowered::stmts(stmts));
        }
        let last_arg = self.last_command_arg(&word, &suffix)?;
        if let Some(error) = self.triggered_parameter_error(&word)? {
            let lowered = Lowered::stmts(parameter_error_stmts(self.subshell_depth > 0, error));
            let lowered = match last_arg {
                Some(arg) => lowered.with_last_arg(arg),
                None => lowered,
            };
            if out.is_empty() && scoped_prefixes.is_empty() && redirs.is_empty() {
                return Ok(lowered);
            }
            let mut stmts = out;
            let body = if redirs.is_empty() {
                lowered.into_stmts()
            } else {
                self.apply_redirections(lowered.into_stmts(), &redirs)?
            };
            if scoped_prefixes.is_empty() {
                stmts.extend(body);
            } else {
                stmts.extend(self.scoped_prefix_stmts(scoped_prefixes, body));
            }
            return Ok(Lowered::stmts(stmts));
        }
        for arg in &suffix {
            if let Some(error) = self.triggered_parameter_error(arg)? {
                let lowered = Lowered::stmts(parameter_error_stmts(self.subshell_depth > 0, error));
                let lowered = match last_arg {
                    Some(arg) => lowered.with_last_arg(arg),
                    None => lowered,
                };
                if out.is_empty() && scoped_prefixes.is_empty() && redirs.is_empty() {
                    return Ok(lowered);
                }
                let mut stmts = out;
                let body = if redirs.is_empty() {
                    lowered.into_stmts()
                } else {
                    self.apply_redirections(lowered.into_stmts(), &redirs)?
                };
                if scoped_prefixes.is_empty() {
                    stmts.extend(body);
                } else {
                    stmts.extend(self.scoped_prefix_stmts(scoped_prefixes, body));
                }
                return Ok(Lowered::stmts(stmts));
            }
        }
        if word_is_bash_command(&word) {
            if let Some(child) = self.child_bash_command(suffix.clone(), &scoped_prefix_values)? {
                let lowered = match last_arg {
                    Some(arg) => child.with_last_arg(arg),
                    None => child,
                };
                if out.is_empty() && scoped_prefixes.is_empty() && redirs.is_empty() {
                    return Ok(lowered);
                }
                let mut stmts = out;
                let body = if redirs.is_empty() {
                    lowered.into_stmts()
                } else {
                    self.apply_redirections(lowered.into_stmts(), &redirs)?
                };
                if scoped_prefixes.is_empty() {
                    stmts.extend(body);
                } else {
                    stmts.extend(self.scoped_prefix_stmts(scoped_prefixes, body));
                }
                return Ok(Lowered::stmts(stmts));
            }
        }

        let lowered = if let Some(n) = name.as_deref() {
            if word_has_unquoted_glob(&word) {
                let mut words = vec![word];
                words.extend(suffix);
                let argv = shell_args_array(self.words_shell_args(words)?);
                Lowered::value(Cmd::status(bash_dynamic_command_status_expr(argv)))
            } else if modifiers.is_empty() {
                if !suppress_alias
                    && let Some(alias_lowered) = self.alias_expansion_call(n, &suffix)?
                {
                    alias_lowered
                } else {
                    self.builtin_or_call(n, &word, suffix, ifs_empty, &modifiers)?
                }
            } else {
                self.builtin_or_call(n, &word, suffix, ifs_empty, &modifiers)?
            }
        } else {
            {
                if !word_has_quoted_part(&word)
                    && self
                        .static_word_text(&word)
                        .is_some_and(|text| text.is_empty())
                {
                    return Ok(Lowered::value(Cmd::status(int(0))));
                }
                if !word_has_quoted_part(&word) {
                    if let Some(text) = self.static_word_text(&word) {
                        let fields = self.static_split_bash_words(&text);
                        if fields.is_empty() {
                            return Ok(Lowered::value(Cmd::status(int(0))));
                        }
                        let mut expanded = fields
                            .iter()
                            .map(|field| bash_quote(field))
                            .collect::<Vec<_>>()
                            .join(" ");
                        for arg in &suffix {
                            expanded.push(' ');
                            expanded.push_str(arg.as_str());
                        }
                        if let Ok(mut pairs) = WashmParser::parse(Rule::command, &expanded) {
                            if let Some(command) = pairs.next() {
                                Lowered::stmts(self.command_stmts(command)?)
                            } else {
                                Lowered::value(Cmd::status(int(0)))
                            }
                        } else {
                            let callee = lit(&text);
                            let args = self.words_exprs(suffix)?;
                            Lowered::value(Cmd::status(call(callee, args)))
                        }
                    } else {
                        let mut words = vec![word];
                        words.extend(suffix);
                        let argv = shell_args_array(self.words_shell_args(words)?);
                        Lowered::value(Cmd::status(bash_dynamic_command_status_expr(argv)))
                    }
                } else {
                    let mut words = vec![word];
                    words.extend(suffix);
                    let argv = shell_args_array(self.words_shell_args(words)?);
                    Lowered::value(Cmd::status(bash_dynamic_command_status_expr(argv)))
                }
            }
        };
        let lowered = match last_arg {
            Some(arg) => lowered.with_last_arg(arg),
            None => lowered,
        };
        let mut stmts = out;
        let body = if redirs.is_empty() {
            lowered.into_stmts()
        } else {
            self.apply_redirections(lowered.into_stmts(), &redirs)?
        };
        let body = self.apply_failglob_guards(body, &failglob_words);
        if scoped_prefixes.is_empty() {
            stmts.extend(body);
        } else {
            stmts.extend(self.scoped_prefix_stmts(scoped_prefixes, body));
        }
        Ok(Lowered::stmts(stmts))
    }

    /// Bash builtins with statement semantics, else a call of `name`.
    fn builtin_or_call(
        &mut self,
        name: &str,
        _word: &Pair<Rule>,
        suffix: Vec<Pair<Rule>>,
        ifs_empty: bool,
        modifiers: &[String],
    ) -> R<Lowered> {
        let bypass_functions = modifiers.iter().any(|m| m == "builtin" || m == "command");
        let force_builtin = modifiers.iter().any(|m| m == "builtin");
        if !bypass_functions && !self.suppress_function_inlining {
            if let Some(lowered) = self.user_function_call(name, &suffix)? {
                return Ok(lowered);
            }
        }
        Ok(match name {
            ":" => Lowered::value(Cmd::status(int(0))),
            "true" => Lowered::value(Cmd::status(int(0))),
            "false" => Lowered::value(Cmd::status(int(1))),
            "return" => {
                let args = self.words_exprs(suffix)?;
                let value = args
                    .into_iter()
                    .next()
                    .map(to_number)
                    .unwrap_or_else(|| ident("__bash_status"));
                let value = if self.function_depth > 0 {
                    let mut exprs = vec![assign_expr(ident("__bash_ret"), value)];
                    exprs.extend(funcname_pop_exprs());
                    exprs.push(ident("__bash_ret"));
                    sequence(exprs)
                } else {
                    value
                };
                Lowered::stmts(vec![Statement::new(StmtKind::Return(Some(value)))])
            }
            "exit" => {
                let args = self.words_exprs(suffix)?;
                let value = args
                    .into_iter()
                    .next()
                    .map(to_number)
                    .unwrap_or_else(|| ident("__bash_status"));
                Lowered::stmts(vec![Statement::new(StmtKind::Exit {
                    status: Some(value),
                })])
            }
            "break" | "continue" => {
                let level = suffix
                    .first()
                    .and_then(literal_word_text)
                    .and_then(|t| t.parse::<u32>().ok());
                Lowered::stmts(vec![Statement::new(if name == "break" {
                    StmtKind::Break(match level {
                        Some(n) if n > 1 => BreakTarget::Level(n),
                        _ => BreakTarget::Implicit,
                    })
                } else {
                    StmtKind::Continue(match level {
                        Some(n) if n > 1 => ContinueTarget::Level(n),
                        _ => ContinueTarget::Implicit,
                    })
                })])
            }
            "local" | "declare" | "typeset" | "readonly" | "export" => {
                Lowered::stmts(self.walk_declaration(name, suffix)?)
            }
            "alias" => self.walk_alias(suffix)?,
            "unalias" => self.walk_unalias(suffix)?,
            "unset" => {
                let mut out = Vec::new();
                let mut functions = false;
                let mut nameref = false;
                let mut status = 0;
                for w in suffix {
                    let Some(t) = self.static_word_text(&w) else {
                        continue;
                    };
                    if t.starts_with('-') {
                        functions = t.contains('f');
                        nameref = t.contains('n');
                        continue;
                    }
                    if functions {
                        self.functions.remove(&t);
                        self.function_bodies.remove(&t);
                        self.function_commands.remove(&t);
                        self.exported_functions.remove(&t);
                        self.traced_functions.remove(&t);
                    } else if nameref {
                        self.namerefs.remove(&t);
                        self.variable_values.remove(&t);
                        out.push(assign_stmt(ident(&t), undefined()));
                    } else if matches!(t.as_str(), "#" | "$" | "!" | "?" | "-" | "_" | "@" | "*") {
                        continue;
                    } else if self.target_text_is_readonly(&t) {
                        status = 1;
                    } else if is_unset_target(&t) {
                        if !t.contains('[') {
                            let resolved =
                                self.resolve_nameref_text(&t).unwrap_or_else(|| t.clone());
                            self.variable_values.remove(&resolved);
                            self.variables.remove(&resolved);
                            self.exported_vars.remove(&resolved);
                        }
                        out.push(assign_stmt(self.name_or_element_target(&t)?, undefined()));
                    } else {
                        status = 1;
                    }
                }
                out.push(assign_stmt(ident("__bash_status"), int(status)));
                Lowered::stmts(out)
            }
            "shift" => {
                let args = self.words_exprs(suffix)?;
                let tmp = self.fresh("__bash_shift");
                let n = args
                    .into_iter()
                    .next()
                    .map(to_number)
                    .unwrap_or_else(|| int(1));
                let in_range = binary(
                    BinOp::And,
                    binary(BinOp::GtEq, ident(&tmp), int(0)),
                    binary(BinOp::LtEq, ident(&tmp), member(ident(BASH_ARGS), "length")),
                );
                Lowered::stmts(vec![
                    let_stmt(&tmp, n),
                    assign_stmt(
                        ident("__bash_status"),
                        ternary(
                            in_range,
                            sequence(vec![
                                assign_expr(
                                    ident(BASH_ARGS),
                                    array_slice_expr(ident(BASH_ARGS), ident(&tmp), None),
                                ),
                                int(0),
                            ]),
                            int(1),
                        ),
                    ),
                ])
            }
            "set" => self.walk_set(suffix)?,
            "wait" => self.walk_wait(suffix)?,
            "shopt" => self.walk_shopt(suffix)?,
            "trap" => self.walk_trap(suffix)?,
            "ulimit" | "umask" | "hash" | "enable" | "disown" => {
                // Shell-state builtins with no model yet; they do not affect output.
                Lowered::stmts(vec![Statement::new(StmtKind::Empty)])
            }
            "exec" => Lowered::value(Cmd::status(int(0))),
            "let" => {
                let mut exprs = Vec::new();
                for w in suffix {
                    let text = word_source_text(&w);
                    exprs.push(self.parse_arith_text(&text)?);
                }
                let last = exprs.pop().unwrap_or_else(|| int(0));
                exprs.push(binary(BinOp::NotEq, last, int(0)));
                Lowered::value(Cmd::boolean(sequence(exprs)))
            }
            "echo" => self.walk_echo(suffix)?,
            "printf" => self.walk_printf(suffix)?,
            "mkdir" => self.walk_mkdir(suffix)?,
            "read" => self.walk_read(suffix, ifs_empty)?,
            "mapfile" | "readarray" => self.walk_mapfile(suffix)?,
            "test" => {
                if suffix.is_empty() {
                    return Ok(Lowered::value(Cmd::status(int(1))));
                }
                if suffix.len() == 1 {
                    return Ok(Lowered::value(Cmd::boolean(binary(
                        BinOp::NotEq,
                        self.word_expr(suffix.into_iter().next().unwrap())?,
                        lit(""),
                    ))));
                }
                if suffix.len() == 2 && literal_word_text(&suffix[0]).as_deref() == Some("!") {
                    return Ok(Lowered::value(Cmd::boolean(binary(
                        BinOp::Eq,
                        self.word_expr(suffix.into_iter().nth(1).unwrap())?,
                        lit(""),
                    ))));
                }
                if suffix.len() == 2 {
                    if let Some(op) = literal_word_text(&suffix[0]) {
                        if op.starts_with('-') && !is_test_unary_operator(&op) {
                            return Ok(Lowered::stmts(vec![
                                bash_stderr_stmt(lit("bash: test: unary operator expected\n")),
                                assign_stmt(ident("__bash_status"), int(2)),
                            ]));
                        }
                    }
                }
                if suffix
                    .iter()
                    .any(|w| literal_word_text(w).as_deref() == Some("]"))
                {
                    return Ok(Lowered::value(Cmd::status(int(2))));
                }
                if let Some(expr) = single_bracket_arith_error_status_expr(&suffix, "test") {
                    return Ok(Lowered::value(Cmd::status(expr)));
                }
                if suffix.iter().any(|w| {
                    word_may_expand_to_multiple_args(w) || word_is_unquoted_expansion_only(w)
                }) {
                    let argv = shell_args_array(self.words_shell_args(suffix)?);
                    return Ok(Lowered::value(Cmd::status(
                        bash_test_status_argv_expr_with_command(argv, "test"),
                    )));
                }
                if is_single_bracket_arith_form(&suffix) {
                    return Ok(Lowered::value(Cmd::status(bash_spawn_status_expr(
                        "test",
                        self.words_exprs(suffix)?,
                    ))));
                }
                let text: Vec<String> = suffix.iter().map(word_source_text).collect();
                let src = format!("[ {} ]", text.join(" "));
                match WashmParser::parse(Rule::test_single_bracket, &src) {
                    Ok(pairs) => {
                        let p = pairs.into_iter().next().ok_or("empty test")?;
                        Lowered::value(Cmd::boolean(self.walk_test_bracket(p)?))
                    }
                    Err(_) => Lowered::value(Cmd::status(bash_spawn_status_expr(
                        "test",
                        self.words_exprs(suffix)?,
                    ))),
                }
            }
            "[" => {
                let mut suffix = suffix;
                let had_closing = suffix.last().and_then(literal_word_text).as_deref() == Some("]");
                if had_closing {
                    suffix.pop();
                } else {
                    return Ok(Lowered::stmts(vec![
                        bash_stderr_stmt(lit("bash: [: missing `]'\n")),
                        assign_stmt(ident("__bash_status"), int(2)),
                    ]));
                }
                if suffix.is_empty() {
                    Lowered::value(Cmd::status(int(1)))
                } else if suffix.len() == 1 {
                    Lowered::value(Cmd::boolean(binary(
                        BinOp::NotEq,
                        self.word_expr(suffix.remove(0))?,
                        lit(""),
                    )))
                } else if suffix.len() == 2 && literal_word_text(&suffix[0]).as_deref() == Some("!")
                {
                    Lowered::value(Cmd::boolean(binary(
                        BinOp::Eq,
                        self.word_expr(suffix.remove(1))?,
                        lit(""),
                    )))
                } else if let Some(expr) = single_bracket_arith_error_status_expr(&suffix, "[") {
                    Lowered::value(Cmd::status(expr))
                } else if suffix.iter().any(|w| {
                    word_may_expand_to_multiple_args(w) || word_is_unquoted_expansion_only(w)
                }) {
                    let argv = shell_args_array(self.words_shell_args(suffix)?);
                    Lowered::value(Cmd::status(bash_test_status_argv_expr(argv)))
                } else if is_single_bracket_arith_form(&suffix) {
                    Lowered::value(Cmd::status(bash_spawn_status_expr(
                        "test",
                        self.words_exprs(suffix)?,
                    )))
                } else {
                    let text: Vec<String> = suffix.iter().map(word_source_text).collect();
                    let src = format!("[ {} ]", text.join(" "));
                    match WashmParser::parse(Rule::test_single_bracket, &src) {
                        Ok(pairs) => {
                            let p = pairs.into_iter().next().ok_or("empty test")?;
                            Lowered::value(Cmd::boolean(self.walk_test_bracket(p)?))
                        }
                        Err(_) => Lowered::value(Cmd::status(bash_spawn_status_expr(
                            "test",
                            self.words_exprs(suffix)?,
                        ))),
                    }
                }
            }
            "source" | "." => self.walk_source(suffix)?,
            "eval" => self.walk_eval(suffix)?,
            "cd" => self.walk_cd(suffix)?,
            "pwd" => Lowered::value(Cmd::status(call_named(
                "printf",
                vec![lit("%s\n"), ident("PWD")],
            ))),
            "env" => self.walk_env(suffix)?,
            "compgen" => self.walk_compgen(suffix)?,
            "type" => self.walk_type(suffix)?,
            _ => {
                if force_builtin {
                    Lowered::value(Cmd::status(int(1)))
                } else {
                    let args = self.words_shell_args(suffix)?;
                    self.external_command(name, args)
                }
            }
        })
    }

    fn walk_wait(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let args = self.words_exprs(suffix)?;
        if args.is_empty() {
            return Ok(Lowered::value(Cmd::status(int(0))));
        }
        let pid = self.fresh("__bash_wait_pid");
        let status = self.fresh("__bash_wait_status");
        Ok(Lowered::value(Cmd::status(sequence(vec![
            assign_expr(ident(&pid), args.into_iter().next().unwrap()),
            assign_expr(ident(&status), index(ident("__bash_jobs"), ident(&pid))),
            ternary(
                binary(BinOp::StrictEq, ident(&status), undefined()),
                int(127),
                ident(&status),
            ),
        ]))))
    }

    fn walk_trap(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        if suffix.len() < 2 {
            return Ok(Lowered::value(Cmd::status(int(0))));
        }
        let action = literal_word_text(&suffix[0]);
        let signal = suffix.last().and_then(literal_word_text);
        if !matches!(signal.as_deref(), Some("EXIT") | Some("0")) {
            return Ok(Lowered::value(Cmd::status(int(0))));
        }
        match action.as_deref() {
            Some("-") => {
                self.exit_trap_body = None;
                Ok(Lowered::value(Cmd::status(int(0))))
            }
            Some(source) => {
                let (stripped, heredocs) = extract_heredocs(source);
                let pairs = WashmParser::parse(Rule::program, &stripped)
                    .map_err(|e| format!("trap command: {e}"))?;
                let mut child = self.clone();
                child.heredocs = heredocs.into();
                let mut body = Vec::new();
                for pair in pairs {
                    if pair.as_rule() == Rule::program {
                        for inner in pair.into_inner() {
                            if inner.as_rule() == Rule::list {
                                body.extend(child.walk_list(inner)?);
                            }
                        }
                    }
                }
                self.counter = child.counter;
                self.exit_trap_body = Some(body);
                Ok(Lowered::value(Cmd::status(int(0))))
            }
            None => Ok(Lowered::value(Cmd::status(int(1)))),
        }
    }

    fn user_function_call(&mut self, name: &str, suffix: &[Pair<Rule>]) -> R<Option<Lowered>> {
        if !is_function_name(name) || !self.functions.contains_key(name) {
            return Ok(None);
        }
        let preserves_static_state = self
            .function_commands
            .get(name)
            .is_some_and(|source| function_source_preserves_static_state(source));
        if let Some(inlined) = self.inline_function_call(name, suffix)? {
            if !preserves_static_state {
                self.variable_values.clear();
                self.array_values.clear();
            }
            return Ok(Some(inlined));
        }
        let fn_name = self.functions.get(name).cloned().unwrap();
        let args = self.words_shell_args(suffix.to_vec())?;
        let (_, state_args) = shell_state_params_args();
        let mut call_args = vec![Argument {
            value: shell_args_array(args),
            name: None,
            by_ref: false,
            spread: false,
        }];
        call_args.extend(state_args.into_iter().map(|value| Argument {
            value,
            name: None,
            by_ref: false,
            spread: false,
        }));
        let call_expr = Expression::new(ExprKind::Call {
            callee: Box::new(ident(&fn_name)),
            args: call_args,
            optional: false,
        });
        if !preserves_static_state {
            self.variable_values.clear();
            self.array_values.clear();
        }
        Ok(Some(Lowered::value(Cmd::status(call_expr))))
    }

    fn alias_expansion_call(&mut self, name: &str, suffix: &[Pair<Rule>]) -> R<Option<Lowered>> {
        if !self.expand_aliases_enabled {
            return Ok(None);
        }
        let Some(replacement) = self.aliases.get(name).cloned() else {
            return Ok(None);
        };
        let mut src = replacement;
        for arg in suffix {
            src.push(' ');
            src.push_str(arg.as_str());
        }
        let pairs = match WashmParser::parse(Rule::command, &src) {
            Ok(pairs) => pairs,
            Err(_) => return Ok(None),
        };
        let mut child = self.clone();
        child.expand_aliases_enabled = false;
        let mut body = Vec::new();
        for pair in pairs {
            body.extend(child.command_stmts(pair)?);
        }
        self.counter = child.counter;
        self.heredocs = child.heredocs;
        Ok(Some(Lowered::stmts(body)))
    }

    fn inline_function_call(&mut self, name: &str, suffix: &[Pair<Rule>]) -> R<Option<Lowered>> {
        let Some(source) = self.function_commands.get(name).cloned() else {
            return Ok(None);
        };
        let nested_call = self.function_depth > 0;
        if source.contains("return")
            || source.contains("<<")
            || source.contains("FUNCNAME")
            || source_has_command_word(&source, name)
            || self.inline_stack.iter().any(|active| active == name)
        {
            return Ok(None);
        }
        if !nested_call && self
            .functions
            .keys()
            .any(|function_name| function_name != name && source_has_command_word(&source, function_name))
        {
            return Ok(None);
        }
        let uses_positionals = source_has_positional_reference(&source);
        let mut positional_values = Vec::new();
        for w in suffix {
            let Some(text) = self.static_word_text(w) else {
                return Ok(None);
            };
            if !word_has_quoted_part(w) && word_text_has_glob_meta(&text) {
                return Ok(None);
            }
            positional_values.push(text);
        }
        if uses_positionals && positional_values.is_empty() && !suffix.is_empty() {
            return Ok(None);
        }
        let pairs = match WashmParser::parse(Rule::command, &source) {
            Ok(pairs) => pairs,
            Err(_) => return Ok(None),
        };
        let mut child = self.clone();
        child.positional_values = positional_values;
        child.function_depth += 1;
        child.dynamic_arith = false;
        child.inline_call_context = true;
        child.inline_stack.push(name.to_string());
        let mut body = Vec::new();
        for pair in pairs {
            body.extend(child.command_stmts(pair)?);
        }
        self.counter = child.counter;
        let mut locals = Vec::new();
        collect_function_scoped_names(&body, &mut locals);
        rewrite_function_scoped_decls(&mut body);
        if !locals.is_empty() {
            let mut wrapped = Vec::new();
            let mut saved = Vec::new();
            for name in locals {
                let tmp = self.fresh("__bash_local_saved");
                wrapped.push(let_stmt(&tmp, ident(&name)));
                saved.push((name, tmp));
            }
            wrapped.extend(body);
            for (name, tmp) in saved.into_iter().rev() {
                wrapped.push(assign_stmt(ident(&name), ident(&tmp)));
            }
            body = wrapped;
        }
        Ok(Some(Lowered::stmts(vec![Statement::new(StmtKind::Block(
            body,
        ))])))
    }

    fn external_command(&mut self, name: &str, args: Vec<ShellArg>) -> Lowered {
        let result = self.fresh("__bash_spawn");
        Lowered::stmts(vec![
            let_stmt(
                &result,
                call_named(
                    "__bash_spawn",
                    vec![
                        lit(name),
                        shell_args_array(args),
                        object(vec![
                            ("input", ident("__bash_stdin")),
                            ("cwd", ident("PWD")),
                        ]),
                    ],
                ),
            ),
            expr_stmt(bash_stdout_write_expr(member(ident(&result), "stdout"))),
            assign_stmt(ident("__bash_status"), spawn_status(ident(&result))),
        ])
    }

    fn child_bash_command(
        &mut self,
        suffix: Vec<Pair<Rule>>,
        prefixes: &[(String, Expression)],
    ) -> R<Option<Lowered>> {
        self.child_bash_command_with_env(suffix, prefixes, false)
    }

    fn child_bash_command_with_env(
        &mut self,
        suffix: Vec<Pair<Rule>>,
        prefixes: &[(String, Expression)],
        clean_env: bool,
    ) -> R<Option<Lowered>> {
        let mut words = Vec::new();
        for w in &suffix {
            let Some(text) = self.static_word_text(w) else {
                return Ok(None);
            };
            words.push(text);
        }
        let Some(pos) = words.iter().position(|w| w == "-c") else {
            return Ok(None);
        };
        let Some(source) = words.get(pos + 1).cloned() else {
            return Ok(Some(Lowered::value(Cmd::status(int(2)))));
        };
        let noexec = words.iter().take(pos).any(|w| {
            w == "-n"
                || (w.starts_with('-')
                    && !w.starts_with("--")
                    && w.chars().skip(1).any(|c| c == 'n'))
        });
        let positional = words
            .iter()
            .skip(pos + 3)
            .map(|s| lit(s))
            .collect::<Vec<_>>();
        let positional_values = words.iter().skip(pos + 3).cloned().collect::<Vec<_>>();

        let (stripped, heredocs) = extract_heredocs(&source);
        if let Some(message) = eval_obvious_syntax_error(&stripped) {
            return Ok(Some(Lowered::value(Cmd::status(bash_c_syntax_error_expr(
                &source, &message,
            )))));
        }
        let pairs = match WashmParser::parse(Rule::program, &stripped) {
            Ok(pairs) => pairs,
            Err(_) => {
                return Ok(Some(Lowered::value(Cmd::status(bash_c_syntax_error_expr(
                    &source,
                    "syntax error",
                )))))
            }
        };
        if noexec {
            return Ok(Some(Lowered::value(Cmd::status(int(0)))));
        }

        let inherited: HashSet<String> = if clean_env {
            HashSet::new()
        } else {
            self.exported_vars.clone()
        };
        let prefix_names: HashSet<String> = prefixes.iter().map(|(name, _)| name.clone()).collect();
        let mut variables = inherited
            .iter()
            .chain(prefix_names.iter())
            .filter(|name| is_name(name))
            .cloned()
            .collect::<BTreeSet<_>>();
        variables.insert("BASH_EXECUTION_STRING".to_string());
        let mut variable_values = HashMap::new();
        variable_values.insert("BASH_EXECUTION_STRING".to_string(), source.clone());
        for name in &variables {
            if let Some(value) = self.variable_values.get(name) {
                variable_values.insert(name.clone(), value.clone());
            }
        }
        for (name, value) in prefixes {
            if let Some(value) = literal_string(value) {
                variable_values.insert(name.clone(), value.to_string());
            } else {
                variable_values.remove(name);
            }
        }

        let mut child = Walker {
            heredocs: heredocs.into(),
            counter: self.counter,
            functions: self.functions.clone(),
            function_bodies: self.function_bodies.clone(),
            function_commands: self.function_commands.clone(),
            variables,
            variable_values,
            array_values: self.array_values.clone(),
            indexed_arrays: self.indexed_arrays.clone(),
            assoc_arrays: self.assoc_arrays.clone(),
            readonly_vars: self.readonly_vars.clone(),
            integer_vars: self.integer_vars.clone(),
            lowercase_vars: self.lowercase_vars.clone(),
            uppercase_vars: self.uppercase_vars.clone(),
            exported_vars: if clean_env {
                prefix_names.clone()
            } else {
                self.exported_vars
                    .iter()
                    .chain(prefix_names.iter())
                    .cloned()
                    .collect()
            },
            exported_functions: self.exported_functions.clone(),
            traced_functions: self.traced_functions.clone(),
            namerefs: self.namerefs.clone(),
            aliases: self.aliases.clone(),
            positional_values,
            brace_expansion_enabled: self.brace_expansion_enabled,
            nullglob_enabled: self.nullglob_enabled,
            dotglob_enabled: self.dotglob_enabled,
            failglob_enabled: self.failglob_enabled,
            globstar_enabled: self.globstar_enabled,
            expand_aliases_enabled: self.expand_aliases_enabled,
            nocaseglob_enabled: self.nocaseglob_enabled,
            nocasematch_enabled: self.nocasematch_enabled,
            patsub_replacement_enabled: self.patsub_replacement_enabled,
            lastpipe_enabled: self.lastpipe_enabled,
            shell_flags: self.shell_flags.clone(),
            subshell_depth: 1,
            arith_stack: Vec::new(),
            dynamic_arith: false,
            function_depth: 0,
            inline_call_context: false,
            inline_stack: Vec::new(),
            suppress_function_inlining: false,
            hoisted_functions: Vec::new(),
            exit_trap_body: None,
        };
        let mut body = Vec::new();
        for pair in pairs {
            if pair.as_rule() == Rule::program {
                for inner in pair.into_inner() {
                    if inner.as_rule() == Rule::list {
                        body.extend(child.walk_list(inner)?);
                    }
                }
            }
        }
        self.counter = child.counter;
        body.push(Statement::new(StmtKind::Return(Some(ident(
            "__bash_status",
        )))));
        let env_value = |name: &str| {
            if name == "BASH_EXECUTION_STRING" {
                lit(&source)
            } else if name == "__bash_last_arg" {
                lit("bash")
            } else if matches!(
                name,
                "__bash_fds"
                    | "__bash_stdin"
                    | "__bash_stdout_null"
                    | "__bash_stdout_to_stderr"
                    | "__bash_stderr_to_stdout"
                    | "__bash_stderr_null"
            ) {
                ident(name)
            } else if let Some((_, value)) = prefixes.iter().find(|(prefix, _)| prefix == name) {
                value.clone()
            } else if inherited.contains(name) {
                ident(name)
            } else {
                undefined()
            }
        };
        let (params, args) = self.shell_child_bindings(&child, false, prefixes, Some(&env_value));
        let mut args = args;
        args[0] = array(positional);
        Ok(Some(Lowered::value(Cmd::status(iife_with_args(
            body, params, args,
        )))))
    }

    fn static_word_text(&self, pair: &Pair<Rule>) -> Option<String> {
        let raw = pair.as_str().trim();
        let unquoted = raw
            .strip_prefix('"')
            .and_then(|s| s.strip_suffix('"'))
            .unwrap_or(raw);
        let defer_to_brace_expansion =
            self.brace_expansion_enabled && raw_may_contain_brace_expansion(unquoted);
        if !defer_to_brace_expansion {
            if let Some(text) = self.static_word_parts_text(pair) {
                return Some(text);
            }
            if let Some(text) = literal_word_text(pair) {
                return Some(text);
            }
        }
        if unquoted == "$BASH" || unquoted == "${BASH}" {
            return Some("bash".to_string());
        }
        if self.function_depth == 0 {
            match unquoted {
                "$#" => return Some(self.positional_values.len().to_string()),
                "$*" | "$@" => return Some(self.positional_values.join(" ")),
                "$-" => return Some(self.shell_flags_text()),
                _ => {}
            }
        }
        if self.function_depth == 0 {
            if let Some(pos) = unquoted
                .strip_prefix('$')
                .and_then(|s| s.parse::<usize>().ok())
            {
                return self.positional_values.get(pos.saturating_sub(1)).cloned();
            }
            if let Some(pos) = unquoted
                .strip_prefix("${")
                .and_then(|s| s.strip_suffix('}'))
                .and_then(|s| s.parse::<usize>().ok())
            {
                return self.positional_values.get(pos.saturating_sub(1)).cloned();
            }
        }
        if !defer_to_brace_expansion {
            if let Some(text) = self.static_word_parts_text(pair) {
                return Some(text);
            }
        }
        if defer_to_brace_expansion {
            return None;
        }
        let mut folded = unquoted.to_string();
        if folded.contains("$((") {
            if let Some(text) = replace_static_arith_expansions(&folded) {
                folded = text;
            }
        }
        if folded.contains("$(") {
            folded = replace_static_command_substitutions(&folded)?;
        }
        if folded.contains('$') {
            let mut text = folded.clone();
            let mut names: Vec<&String> = self.variable_values.keys().collect();
            names.sort_by_key(|name| std::cmp::Reverse(name.len()));
            for name in names {
                let value = self.variable_values.get(name)?;
                text = text.replace(&format!("${{{name}}}"), value);
                text = text.replace(&format!("${name}"), value);
            }
            if text.contains("$((") {
                text = replace_static_arith_expansions(&text)?;
            }
            if !text.contains('$') {
                return Some(text);
            }
        }
        let name = whole_unquoted_param_name(&folded)?;
        let resolved = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        self.variable_values.get(&resolved).cloned()
    }

    fn static_word_parts_text(&self, pair: &Pair<Rule>) -> Option<String> {
        if !matches!(pair.as_rule(), Rule::word | Rule::assignment_value_word) {
            return None;
        }
        let mut out = String::new();
        for part in pair.clone().into_inner() {
            out.push_str(&self.static_part_text(&part)?);
        }
        Some(out)
    }

    fn static_part_text(&self, pair: &Pair<Rule>) -> Option<String> {
        Some(match pair.as_rule() {
            Rule::bare_word | Rule::assignment_bare_word => unescape_bare(pair.as_str()),
            Rule::close_brace_tail => "}".to_string(),
            Rule::single_quoted_string => {
                let s = pair.as_str();
                s[1..s.len() - 1].to_string()
            }
            Rule::ansi_c_quoting => {
                let s = pair.as_str();
                decode_ansi_c(&s[2..s.len() - 1])
            }
            Rule::quoted_string | Rule::locale_quoting => {
                let mut out = String::new();
                for q in pair.clone().into_inner() {
                    match q.as_rule() {
                        Rule::dq_text => out.push_str(q.as_str()),
                        Rule::dq_escape => {
                            let c = q.as_str().chars().nth(1).unwrap_or('\\');
                            out.push_str(&match c {
                                '$' | '`' | '"' | '\\' => c.to_string(),
                                '\n' => String::new(),
                                _ => format!("\\{c}"),
                            });
                        }
                        _ => out.push_str(&self.static_part_text(&q)?),
                    }
                }
                out
            }
            Rule::simple_param => {
                let inner = pair.clone().into_inner().next()?;
                match inner.as_rule() {
                    Rule::name => self.static_name_value(inner.as_str())?,
                    Rule::positional_digit => {
                        let pos = inner.as_str().parse::<usize>().ok()?;
                        self.positional_values.get(pos.saturating_sub(1)).cloned()?
                    }
                    Rule::special_param => match inner.as_str() {
                        "#" if self.function_depth == 0 => self.positional_values.len().to_string(),
                        "*" | "@" if self.function_depth == 0 => self.positional_values.join(" "),
                        "-" => self.shell_flags_text(),
                        _ => return None,
                    },
                    _ => return None,
                }
            }
            Rule::braced_param => {
                if pair.as_str().starts_with("${!") {
                    return None;
                }
                if let Some((name, op)) = static_transform_braced(pair.as_str()) {
                    let resolved = self
                        .resolve_nameref_text(name)
                        .unwrap_or_else(|| name.to_string());
                    return Some(match op {
                        "U" => self.static_name_value(&resolved)?.to_uppercase(),
                        "L" => self.static_name_value(&resolved)?.to_lowercase(),
                        "u" => {
                            let value = self.static_name_value(&resolved)?;
                            let mut chars = value.chars();
                            match chars.next() {
                                Some(first) => {
                                    format!("{}{}", first.to_uppercase(), chars.as_str())
                                }
                                None => String::new(),
                            }
                        }
                        "E" | "P" => decode_ansi_c(&self.static_name_value(&resolved)?),
                        "Q" => {
                            if !self.variable_is_set(&resolved) {
                                String::new()
                            } else {
                                bash_quote(&self.static_name_value(&resolved)?)
                            }
                        }
                        "a" => self.attribute_letters(&resolved),
                        _ => return None,
                    });
                }
                let name = simple_braced_param_name(pair.as_str())?;
                self.static_name_value(name)?
            }
            Rule::tilde_expansion => return None,
            Rule::arithmetic_expansion => {
                if self.dynamic_arith {
                    return None;
                }
                let inner = arithmetic_expansion_inner(pair.as_str())?;
                if inner.contains('$') {
                    let expanded = self.static_arith_parameter_substitution(inner)?;
                    return Some(self.eval_static_arith_text(&expanded)?.to_string());
                }
                if inner.contains('$')
                    || inner
                        .chars()
                        .any(|ch| ch == '_' || ch.is_ascii_alphabetic())
                {
                    return None;
                }
                self.eval_static_arith_text(inner)?.to_string()
            }
            Rule::dollar_paren_subst => return None,
            _ => return None,
        })
    }

    fn static_name_value(&self, name: &str) -> Option<String> {
        if name == "BASH" {
            return Some("bash".to_string());
        }
        if name == "BASH_SUBSHELL" {
            return Some(self.subshell_depth.to_string());
        }
        let resolved = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        self.variable_values.get(&resolved).cloned()
    }

    fn triggered_parameter_error(&mut self, pair: &Pair<Rule>) -> R<Option<(String, Expression)>> {
        if pair.as_rule() == Rule::error_value {
            return self.triggered_error_expansion(pair);
        }
        if pair.as_rule() == Rule::assign_default {
            return Ok(self.triggered_assign_default_error(pair));
        }
        if pair.as_rule() == Rule::arithmetic_expansion {
            if let Some(inner) = arithmetic_expansion_inner(pair.as_str()) {
                let expanded = if inner.contains('$') {
                    self.static_arith_parameter_substitution(inner)
                } else {
                    None
                };
                let text = expanded.as_deref().unwrap_or(inner);
                if let Some(message) = bash_arithmetic_error_message(text) {
                    return Ok(Some(("bash".to_string(), lit(message))));
                }
            }
        }
        if matches!(
            pair.as_rule(),
            Rule::dollar_paren_subst
                | Rule::backtick_subst
                | Rule::brace_command_subst
                | Rule::brace_command_subst_inner
        ) {
            return Ok(None);
        }
        for inner in pair.clone().into_inner() {
            if let Some(error) = self.triggered_parameter_error(&inner)? {
                return Ok(Some(error));
            }
        }
        Ok(None)
    }

    fn triggered_assign_default_error(
        &self,
        pair: &Pair<Rule>,
    ) -> Option<(String, Expression)> {
        let base = pair
            .clone()
            .into_inner()
            .find(|inner| inner.as_rule() == Rule::param_base)?;
        let name = base.as_str().to_string();
        if is_readonly_parameter_target(&name) {
            Some((name, lit("cannot assign in this way")))
        } else {
            None
        }
    }

    fn triggered_error_expansion(&mut self, pair: &Pair<Rule>) -> R<Option<(String, Expression)>> {
        let mut base = None;
        let mut op = String::new();
        let mut word = None;
        for inner in pair.clone().into_inner() {
            match inner.as_rule() {
                Rule::param_base => base = Some(inner),
                Rule::error_op => op = inner.as_str().to_string(),
                Rule::brace_word => word = Some(inner),
                _ => {}
            }
        }
        let Some(base) = base else {
            return Ok(None);
        };
        let Some((name, is_set, value)) = self.static_param_error_target(&base) else {
            return Ok(None);
        };
        let treats_empty_as_error = op.starts_with(':');
        if is_set && (!treats_empty_as_error || !value.is_empty()) {
            return Ok(None);
        }
        let message = match word {
            Some(word) => self.brace_word_expr(word)?,
            None => lit(&format!("{name}: parameter null or not set")),
        };
        Ok(Some((name, message)))
    }

    fn static_param_error_target(&self, pair: &Pair<Rule>) -> Option<(String, bool, String)> {
        let mut name = None;
        let mut subscript = None;
        for inner in pair.clone().into_inner() {
            match inner.as_rule() {
                Rule::name => name = Some(inner.as_str().to_string()),
                Rule::subscript => subscript = Some(inner.as_str().to_string()),
                Rule::positional_param => {
                    let pos = inner.as_str().parse::<usize>().ok()?;
                    let value = self.positional_values.get(pos.saturating_sub(1)).cloned();
                    return Some((
                        inner.as_str().to_string(),
                        value.is_some(),
                        value.unwrap_or_default(),
                    ));
                }
                Rule::special_param => {
                    let value = match inner.as_str() {
                        "#" => self.positional_values.len().to_string(),
                        "?" => "0".to_string(),
                        "$" => "0".to_string(),
                        "*" | "@" => self.positional_values.join(" "),
                        _ => String::new(),
                    };
                    return Some((inner.as_str().to_string(), true, value));
                }
                _ => {}
            }
        }
        let name = name?;
        let resolved = self
            .resolve_nameref_text(&name)
            .unwrap_or_else(|| name.clone());
        let display = subscript
            .as_ref()
            .map(|sub| format!("{name}[{sub}]"))
            .unwrap_or_else(|| name.clone());
        let Some(subscript) = subscript else {
            if let Some(value) = self.variable_values.get(&resolved) {
                return Some((display, true, value.clone()));
            }
            if let Some(entries) = self.array_values.get(&resolved) {
                return Some((
                    display,
                    true,
                    entries
                        .first()
                        .map(|(_, value)| value.clone())
                        .unwrap_or_default(),
                ));
            }
            return Some((display, self.variables.contains(&resolved), String::new()));
        };
        let key = self.static_subscript_key(&resolved, &subscript)?;
        let value = self
            .array_values
            .get(&resolved)
            .and_then(|entries| entries.iter().find(|(k, _)| k == &key))
            .map(|(_, value)| value.clone());
        Some((display, value.is_some(), value.unwrap_or_default()))
    }

    fn static_subscript_key(&self, array: &str, subscript: &str) -> Option<String> {
        let text = subscript.trim();
        if text == "@" || text == "*" {
            return Some(text.to_string());
        }
        if (text.starts_with('"') && text.ends_with('"'))
            || (text.starts_with('\'') && text.ends_with('\''))
        {
            return Some(text[1..text.len() - 1].to_string());
        }
        if let Some(name) = whole_unquoted_param_name(text) {
            return self.static_name_value(name);
        }
        if self.assoc_arrays.contains(array) {
            return Some(unescape_bare(text));
        }
        self.eval_static_arith_text(text)
            .map(|n| n.to_string())
            .or_else(|| text.parse::<i64>().ok().map(|n| n.to_string()))
            .or_else(|| Some(unescape_bare(text)))
    }

    fn eval_static_arith_text(&self, input: &str) -> Option<i64> {
        let mut out = String::new();
        let mut chars = input.char_indices().peekable();
        while let Some((_, ch)) = chars.next() {
            if ch == '$' {
                let mut name = String::new();
                if let Some((_, next)) = chars.peek().copied() {
                    if next == '{' {
                        chars.next();
                        while let Some((_, next)) = chars.peek().copied() {
                            chars.next();
                            if next == '}' {
                                break;
                            }
                            name.push(next);
                        }
                    } else if next.is_ascii_alphabetic() || next == '_' {
                        while let Some((_, next)) = chars.peek().copied() {
                            if next.is_ascii_alphanumeric() || next == '_' {
                                name.push(next);
                                chars.next();
                            } else {
                                break;
                            }
                        }
                    } else {
                        out.push(ch);
                        continue;
                    }
                }
                let value = self.static_name_value(&name)?;
                out.push_str(value.trim());
            } else if ch.is_ascii_alphabetic() || ch == '_' {
                let mut name = String::new();
                name.push(ch);
                while let Some((_, next)) = chars.peek().copied() {
                    if next.is_ascii_alphanumeric() || next == '_' {
                        name.push(next);
                        chars.next();
                    } else {
                        break;
                    }
                }
                let value = self.static_name_value(&name)?;
                let value = value.trim();
                if value.is_empty() {
                    out.push('0');
                } else {
                    out.push('(');
                    out.push_str(value);
                    out.push(')');
                }
            } else {
                out.push(ch);
            }
        }
        eval_const_arith(&out)
    }

    fn shell_child_bindings(
        &self,
        child: &Walker,
        inherit_all: bool,
        prefixes: &[(String, Expression)],
        value_for: Option<&dyn Fn(&str) -> Expression>,
    ) -> (Vec<Param>, Vec<Expression>) {
        let mut names = if inherit_all {
            self.variables
                .iter()
                .chain(child.variables.iter())
                .cloned()
                .collect::<BTreeSet<_>>()
        } else {
            self.variables
                .iter()
                .chain(child.variables.iter())
                .chain(self.exported_vars.iter())
                .cloned()
                .collect::<BTreeSet<_>>()
        };
        names.extend(prefixes.iter().map(|(name, _)| name.clone()));
        names.insert("__bash_last_arg".to_string());
        names.insert("__bash_fds".to_string());
        names.insert("__bash_stdin".to_string());
        names.insert("__bash_abort_line".to_string());
        names.insert("__bash_stdout_null".to_string());
        names.insert("__bash_stdout_to_stderr".to_string());
        names.insert("__bash_stderr_to_stdout".to_string());
        names.insert("__bash_stderr_null".to_string());
        names.insert("__bash_flags".to_string());
        let mut params = vec![args_param()];
        let mut args = vec![ident(BASH_ARGS)];
        for name in names {
            if !is_name(&name) || name == BASH_ARGS {
                continue;
            }
            params.push(param_named(&name));
            let value = if let Some(value_for) = value_for {
                value_for(&name)
            } else if inherit_all && name == "BASH_SUBSHELL" {
                binary(BinOp::Add, ident("BASH_SUBSHELL"), int(1))
            } else if inherit_all
                && (self.variables.contains(&name)
                    || matches!(
                        name.as_str(),
                        "__bash_last_arg"
                            | "__bash_fds"
                            | "__bash_stdin"
                            | "__bash_abort_line"
                            | "__bash_stdout_null"
                            | "__bash_stdout_to_stderr"
                            | "__bash_stderr_to_stdout"
                            | "__bash_stderr_null"
                            | "__bash_flags"
                    ))
            {
                ident(&name)
            } else {
                undefined()
            };
            args.push(value);
        }
        (params, args)
    }

    fn catch_subshell_exit_throw(&mut self, body: Vec<Statement>) -> Vec<Statement> {
        let err = self.fresh("__bash_subshell_exit");
        vec![Statement::new(StmtKind::Try {
            body,
            catches: vec![CatchClause {
                types: Vec::new(),
                var_name: Some(err.clone()),
                stack_var: None,
                body: vec![Statement::new(StmtKind::If {
                    cond: binary(
                        BinOp::StrictEq,
                        member(ident(&err), "__bash_exit"),
                        Expression::bool(true),
                    ),
                    then_body: vec![Statement::new(StmtKind::Return(Some(member(
                        ident(&err),
                        "status",
                    ))))],
                    elifs: Vec::new(),
                    else_body: Some(vec![Statement::new(StmtKind::Throw {
                        expr: Some(ident(&err)),
                        cause: None,
                    })]),
                })],
                when_clause: None,
            }],
            else_body: None,
            finally: None,
        })]
    }

    fn scoped_prefix_stmts(
        &mut self,
        prefixes: Vec<(String, Statement)>,
        body: Vec<Statement>,
    ) -> Vec<Statement> {
        let mut saved = Vec::new();
        let mut out = Vec::new();
        for (name, stmt) in &prefixes {
            let tmp = self.fresh("__bash_saved");
            out.push(let_stmt(&tmp, ident(name)));
            out.push(stmt.clone());
            saved.push((name.clone(), tmp));
        }
        out.extend(body);
        for (name, tmp) in saved.into_iter().rev() {
            out.push(assign_stmt(
                ident(&name),
                ternary(
                    binary(BinOp::StrictEq, ident(&tmp), undefined()),
                    lit(""),
                    ident(&tmp),
                ),
            ));
        }
        out
    }

    fn walk_cd(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let target = match suffix.first() {
            Some(w) => self.word_expr(w.clone())?,
            None => param_value(ident("HOME")),
        };
        self.record_variable("PWD");
        self.record_variable("OLDPWD");
        self.variable_values.remove("PWD");
        self.variable_values.remove("OLDPWD");
        Ok(Lowered::stmts(vec![
            assign_stmt(ident("OLDPWD"), ident("PWD")),
            assign_stmt(ident("PWD"), target),
            assign_stmt(ident("__bash_status"), int(0)),
        ]))
    }

    fn walk_env(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut clean_env = false;
        let mut prefixes: Vec<(String, Expression)> = Vec::new();
        let mut idx = 0usize;
        while let Some(w) = suffix.get(idx) {
            let Some(text) = self.static_word_text(w) else {
                if let Some((name, value)) = self.env_prefix_assignment(w.clone())? {
                    prefixes.push((name, value));
                    idx += 1;
                    continue;
                }
                break;
            };
            if text == "-i" || text == "--ignore-environment" {
                clean_env = true;
                idx += 1;
                continue;
            }
            if let Some((name, value)) = text.split_once('=') {
                if is_name(name) {
                    prefixes.push((name.to_string(), lit(value)));
                    idx += 1;
                    continue;
                }
            }
            break;
        }
        let Some(command) = suffix.get(idx) else {
            return Ok(Lowered::value(Cmd::status(int(0))));
        };
        if self.static_word_text(command).as_deref() == Some("bash") {
            let rest = suffix.clone().into_iter().skip(idx + 1).collect();
            if let Some(child) = self.child_bash_command_with_env(rest, &prefixes, clean_env)? {
                return Ok(child);
            }
        }
        let mut args = Vec::new();
        for (_, value) in prefixes {
            args.push(ShellArg {
                value,
                spread: false,
            });
        }
        args.extend(self.words_shell_args(suffix.into_iter().skip(idx).collect())?);
        Ok(self.external_command("env", args))
    }

    fn env_prefix_assignment(&mut self, pair: Pair<Rule>) -> R<Option<(String, Expression)>> {
        if pair.as_rule() == Rule::assignment_word {
            let (target, _, value) = self.assignment_parts(pair)?;
            if let AssignTarget::Name(name) = target {
                if is_name(&name) {
                    return Ok(Some((name, value)));
                }
            }
            return Ok(None);
        }
        let Some(text) = literal_word_text(&pair) else {
            return Ok(None);
        };
        let Some((name, value)) = text.split_once('=') else {
            return Ok(None);
        };
        Ok(is_name(name).then(|| (name.to_string(), lit(value))))
    }

    fn walk_compgen(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut variable_names = false;
        let mut prefix = String::new();
        let mut it = suffix.iter();
        while let Some(w) = it.next() {
            let Some(text) = literal_word_text(w) else {
                continue;
            };
            if text == "-v" {
                variable_names = true;
                if let Some(next) = it.next().and_then(literal_word_text) {
                    prefix = next;
                }
                continue;
            }
        }
        if !variable_names {
            return Ok(Lowered::value(Cmd::status(int(1))));
        }
        let mut names: Vec<String> = self
            .variables
            .iter()
            .filter(|name| name.starts_with(&prefix))
            .cloned()
            .collect();
        names.sort();
        let text = names.join("\n");
        Ok(Lowered::value(Cmd::status(call_named(
            "printf",
            vec![lit("%s"), lit(&text)],
        ))))
    }

    fn walk_type(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut type_only = false;
        let mut args = Vec::new();
        for w in suffix {
            if let Some(text) = literal_word_text(&w) {
                if text.starts_with('-') {
                    if text.contains('t') {
                        type_only = true;
                    }
                    continue;
                }
            }
            let saved = self.suppress_function_inlining;
            self.suppress_function_inlining = true;
            let expanded = self.expand_word(w);
            self.suppress_function_inlining = saved;
            args.extend(expanded?);
        }
        if !type_only {
            return Ok(Lowered::value(Cmd::status(int(1))));
        }
        let mut out = Vec::new();
        let mut status_parts = Vec::new();
        for arg in args {
            let kind = self.fresh("__bash_type_kind");
            let kind_expr = literal_string(&arg)
                .map(|name| lit(&self.bash_command_kind_static(name)))
                .unwrap_or_else(|| self.bash_command_kind_expr(arg));
            out.push(let_stmt(&kind, kind_expr));
            out.push(Statement::new(StmtKind::If {
                cond: binary(BinOp::StrictNotEq, ident(&kind), lit("")),
                then_body: vec![expr_stmt(call_named(
                    "printf",
                    vec![lit("%s\n"), ident(&kind)],
                ))],
                elifs: Vec::new(),
                else_body: None,
            }));
            status_parts.push(binary(BinOp::StrictEq, ident(&kind), lit("")));
        }
        let status = status_parts
            .into_iter()
            .reduce(|left, right| binary(BinOp::Or, left, right))
            .unwrap_or_else(|| Expression::bool(false));
        out.push(assign_stmt(
            ident("__bash_status"),
            ternary(status, int(1), int(0)),
        ));
        Ok(Lowered::stmts(out))
    }

    fn walk_alias(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut status = 0;
        for w in suffix {
            let raw = w.as_str().trim();
            let Some((name, value)) = raw.split_once('=') else {
                if !self.aliases.contains_key(raw) {
                    status = 1;
                }
                continue;
            };
            if !is_function_name(name) {
                status = 1;
                continue;
            }
            let value = shell_static_words(value)
                .map(|parts| parts.join(" "))
                .unwrap_or_else(|| value.to_string());
            self.aliases.insert(name.to_string(), value);
        }
        Ok(Lowered::value(Cmd::status(int(status))))
    }

    fn walk_unalias(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut status = 0;
        for w in suffix {
            let Some(text) = self.static_word_text(&w) else {
                status = 1;
                continue;
            };
            if text == "-a" {
                self.aliases.clear();
            } else if self.aliases.remove(&text).is_none() {
                status = 1;
            }
        }
        Ok(Lowered::value(Cmd::status(int(status))))
    }

    fn bash_command_kind_expr(&self, name: Expression) -> Expression {
        let mut functions: Vec<String> = self.functions.keys().cloned().collect();
        functions.sort();
        ternary(
            string_equals_any_expr(name.clone(), BASH_KEYWORD_NAMES),
            lit("keyword"),
            ternary(
                string_equals_any_owned_expr(name.clone(), &functions),
                lit("function"),
                ternary(
                    string_equals_any_expr(name, BASH_BUILTIN_NAMES),
                    lit("builtin"),
                    lit(""),
                ),
            ),
        )
    }

    fn bash_command_kind_static(&self, name: &str) -> String {
        if BASH_KEYWORD_NAMES.contains(&name) {
            "keyword".to_string()
        } else if self.functions.contains_key(name) {
            "function".to_string()
        } else if BASH_BUILTIN_NAMES.contains(&name) {
            "builtin".to_string()
        } else {
            String::new()
        }
    }

    fn walk_source(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let Some(first) = suffix.first() else {
            return Ok(Lowered::value(Cmd::status(int(1))));
        };

        if let Some(src) = self.static_source_text(first) {
            return self.walk_sourced_text(&src);
        }

        let mut args = self.words_exprs(suffix)?;
        let source = args.remove(0);
        let text = self.fresh("__bash_source_text");
        Ok(Lowered::value(Cmd::status(sequence(vec![
            assign_expr(ident(&text), self.read_file_or_fd_expr(source)),
            ternary(
                binary(BinOp::NotEq, ident("__bash_status"), int(0)),
                ident("__bash_status"),
                call_named("__vybe_eval", vec![ident(&text), lit("washm")]),
            ),
        ]))))
    }

    fn static_source_text(&self, pair: &Pair<Rule>) -> Option<String> {
        if pair.as_rule() != Rule::word {
            return None;
        }
        let raw = pair.as_str().trim();
        let inner = raw.strip_prefix("<(")?.strip_suffix(')')?.trim();
        static_command_output(inner)
    }

    fn static_split_bash_words(&self, value: &str) -> Vec<String> {
        if value.is_empty() {
            return Vec::new();
        }
        let ifs = self.variable_values.get("IFS").map(String::as_str);
        let ifs = match ifs {
            Some("") => return vec![value.to_string()],
            Some(ifs) => ifs,
            None => " \t\n",
        };
        let ifs_whitespace = |ch: char| ifs.contains(ch) && ch.is_whitespace();
        let ifs_non_whitespace = |ch: char| ifs.contains(ch) && !ch.is_whitespace();

        if !ifs.chars().any(|ch| !ch.is_whitespace()) {
            let trimmed = value.trim_matches(ifs_whitespace);
            if trimmed.is_empty() {
                return Vec::new();
            }
            return trimmed
                .split(ifs_whitespace)
                .filter(|part| !part.is_empty())
                .map(ToString::to_string)
                .collect();
        }

        let chars = value.chars().collect::<Vec<_>>();
        let mut out = Vec::new();
        let mut field = String::new();
        let mut i = 0;

        while i < chars.len() && ifs_whitespace(chars[i]) {
            i += 1;
        }

        while i < chars.len() {
            let ch = chars[i];
            if ifs_non_whitespace(ch) {
                out.push(std::mem::take(&mut field));
                i += 1;
                while i < chars.len() && ifs_whitespace(chars[i]) {
                    i += 1;
                }
            } else if ifs_whitespace(ch) {
                let mut j = i;
                while j < chars.len() && ifs_whitespace(chars[j]) {
                    j += 1;
                }
                if j < chars.len() && ifs_non_whitespace(chars[j]) {
                    i = j;
                    continue;
                }
                if !field.is_empty() {
                    out.push(std::mem::take(&mut field));
                }
                i += 1;
                while i < chars.len() && ifs_whitespace(chars[i]) {
                    i += 1;
                }
            } else {
                field.push(ch);
                i += 1;
            }
        }

        if !field.is_empty() {
            out.push(field);
        }
        out
    }

    fn walk_sourced_text(&mut self, src: &str) -> R<Lowered> {
        if let Some(message) = eval_obvious_syntax_error(src) {
            return Ok(Lowered::value(Cmd::status(sequence(vec![
                bash_stderr_write_expr(lit(&format!("{message}\n"))),
                int(2),
            ]))));
        }
        let (stripped, heredocs) = extract_heredocs(src);
        match WashmParser::parse(Rule::program, &stripped) {
            Ok(pairs) => {
                let saved_heredocs = std::mem::replace(&mut self.heredocs, heredocs.into());
                let mut body = Vec::new();
                let result = (|| -> R<()> {
                    for pair in pairs {
                        if pair.as_rule() == Rule::program {
                            for inner in pair.into_inner() {
                                if inner.as_rule() == Rule::list {
                                    body.extend(self.walk_list(inner)?);
                                }
                            }
                        }
                    }
                    Ok(())
                })();
                self.heredocs = saved_heredocs;
                result?;
                if body.is_empty() {
                    body.push(assign_stmt(ident("__bash_status"), int(0)));
                }
                Ok(Lowered::stmts(body))
            }
            Err(_) => Ok(Lowered::value(Cmd::status(sequence(vec![
                bash_stderr_write_expr(lit("syntax error\n")),
                int(2),
            ])))),
        }
    }

    fn walk_eval(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        if let Some(lowered) = self.walk_static_eval_command(&suffix)? {
            return Ok(lowered);
        }
        let mut texts = Vec::new();
        for w in &suffix {
            let Some(t) = self.static_word_text(w) else {
                let source = array_join(array(self.words_exprs(suffix)?), lit(" "));
                return Ok(Lowered::value(Cmd::status(call_named(
                    "__vybe_eval",
                    vec![source, lit("washm")],
                ))));
            };
            texts.push(t);
        }
        let src = texts.join(" ");
        if let Some(message) = eval_obvious_expansion_error(&src, &self.variable_values) {
            return Ok(Lowered::stmts(vec![
                bash_stderr_stmt(lit(&format!("{message}\n"))),
                assign_stmt(ident("__bash_status"), int(1)),
            ]));
        }
        if let Some(message) = eval_obvious_syntax_error(&src) {
            return Ok(Lowered::value(Cmd::status(sequence(vec![
                bash_stderr_write_expr(lit(&format!("{message}\n"))),
                int(2),
            ]))));
        }
        let (stripped, heredocs) = extract_heredocs(&src);
        match WashmParser::parse(Rule::program, &stripped) {
            Ok(pairs) => {
                let saved_heredocs = std::mem::replace(&mut self.heredocs, heredocs.into());
                let mut body = Vec::new();
                for pair in pairs {
                    if pair.as_rule() == Rule::program {
                        for inner in pair.into_inner() {
                            if inner.as_rule() == Rule::list {
                                body.extend(self.walk_list(inner)?);
                            }
                        }
                    }
                }
                self.heredocs = saved_heredocs;
                Ok(Lowered::stmts(body))
            }
            Err(_) => Ok(Lowered::value(Cmd::status(sequence(vec![
                bash_stderr_write_expr(lit("syntax error\n")),
                int(2),
            ])))),
        }
    }

    fn walk_static_eval_command(&mut self, suffix: &[Pair<Rule>]) -> R<Option<Lowered>> {
        let Some(first) = suffix.first() else {
            return Ok(None);
        };
        let Some(name) = self.static_word_text(first) else {
            return Ok(None);
        };
        if name != "printf" || suffix.len() < 2 {
            return Ok(None);
        }
        let mut it = suffix.iter().skip(1);
        let fmt_word = it.next().unwrap().clone();
        let mut fmt = self.eval_second_pass_word_expr(fmt_word)?;
        let mut fmt_text = None;
        if let ExprKind::Lit(Literal::Str(s)) = &fmt.kind {
            fmt_text = Some(decode_printf_escapes(s));
        }
        if let Some(s) = &fmt_text {
            fmt = lit(s);
        }
        let mut shell_args = Vec::new();
        for arg in it {
            shell_args.push(ShellArg {
                value: self.eval_second_pass_word_expr(arg.clone())?,
                spread: false,
            });
        }
        let sprintf = self.bash_printf_expr(fmt, fmt_text.as_deref(), shell_args);
        Ok(Some(Lowered::value(Cmd::status(bash_stdout_status_expr(
            sprintf,
        )))))
    }

    fn eval_second_pass_word_expr(&mut self, word: Pair<Rule>) -> R<Expression> {
        if let Some(expr) = self.eval_second_pass_tilde_expr(&word) {
            return Ok(expr);
        }
        self.word_expr(word)
    }

    fn eval_second_pass_tilde_expr(&mut self, word: &Pair<Rule>) -> Option<Expression> {
        let raw = word.as_str().trim();
        let unquoted = raw
            .strip_prefix('"')
            .and_then(|s| s.strip_suffix('"'))
            .unwrap_or(raw);
        let name = unquoted.strip_prefix("~$").filter(|name| is_name(name))?;
        let value = param_value(ident(name));
        Some(ternary(
            binary(
                BinOp::StrictEq,
                bash_env_value_expr("USER", lit("")),
                value.clone(),
            ),
            param_value(ident("HOME")),
            parts_to_expr(vec![Part::Text("~".to_string()), Part::Expr(value)]),
        ))
    }

    /// `set -- a b` and `set a b` replace positional parameters; supported
    /// option flags update the shell flag set used by later normalization.
    fn walk_set(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut it = suffix.into_iter().peekable();
        let mut positional = false;
        let mut out = Vec::new();
        while let Some(w) = it.peek() {
            let Some(t) = literal_word_text(w) else { break };
            if t == "--" {
                it.next();
                positional = true;
                break;
            }
            if t.starts_with('-') || t.starts_with('+') {
                it.next();
                let enable = t.starts_with('-');
                if t == "+B" {
                    self.brace_expansion_enabled = false;
                    self.shell_flags.remove(&'B');
                } else if t == "-B" {
                    self.brace_expansion_enabled = true;
                    self.shell_flags.insert('B');
                } else if t == "-f" {
                    self.shell_flags.insert('f');
                } else if t == "+f" {
                    self.shell_flags.remove(&'f');
                } else if t == "-u" {
                    self.shell_flags.insert('u');
                } else if t == "+u" {
                    self.shell_flags.remove(&'u');
                } else if t == "-a" {
                    self.shell_flags.insert('a');
                } else if t == "+a" {
                    self.shell_flags.remove(&'a');
                } else if t == "-e" {
                    self.shell_flags.insert('e');
                } else if t == "+e" {
                    self.shell_flags.remove(&'e');
                }
                if t == "-o" || t == "+o" {
                    if let Some(opt) = it.next().and_then(|w| literal_word_text(&w)) {
                        let flag = match opt.as_str() {
                            "errexit" => Some('e'),
                            "nounset" => Some('u'),
                            "noglob" => Some('f'),
                            "braceexpand" => Some('B'),
                            "hashall" => Some('h'),
                            "pipefail" => Some('P'),
                            _ => None,
                        };
                        if let Some(flag) = flag {
                            if enable { self.shell_flags.insert(flag); } else { self.shell_flags.remove(&flag); }
                        }
                    }
                }
                out.push(assign_stmt(
                    ident("__bash_flags"),
                    lit(&self.shell_flags_text()),
                ));
                continue;
            }
            break;
        }
        let rest: Vec<Pair<Rule>> = it.collect();
        if !positional && rest.is_empty() {
            out.push(assign_stmt(ident("__bash_status"), int(0)));
            return Ok(Lowered::stmts(out));
        }
        let words = self.words_shell_args(rest)?;
        if words
            .iter()
            .all(|arg| !arg.spread && literal_string(&arg.value).is_some())
        {
            self.positional_values = words
                .iter()
                .filter_map(|arg| literal_string(&arg.value).map(str::to_string))
                .collect();
        } else {
            self.positional_values.clear();
        }
        out.push(assign_stmt(ident(BASH_ARGS), shell_args_array(words)));
        out.push(assign_stmt(ident("__bash_status"), int(0)));
        Ok(Lowered::stmts(out))
    }

    fn walk_shopt(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut mode: Option<bool> = None;
        let mut status = 0;
        for w in suffix {
            let Some(text) = self.static_word_text(&w) else {
                status = 1;
                continue;
            };
            match text.as_str() {
                "-s" => mode = Some(true),
                "-u" => mode = Some(false),
                "nullglob" => {
                    if let Some(enabled) = mode {
                        self.nullglob_enabled = enabled;
                    }
                }
                "dotglob" => {
                    if let Some(enabled) = mode {
                        self.dotglob_enabled = enabled;
                    }
                }
                "failglob" => {
                    if let Some(enabled) = mode {
                        self.failglob_enabled = enabled;
                    }
                }
                "globstar" => {
                    if let Some(enabled) = mode {
                        self.globstar_enabled = enabled;
                    }
                }
                "expand_aliases" => {
                    if let Some(enabled) = mode {
                        self.expand_aliases_enabled = enabled;
                    }
                }
                "nocaseglob" => {
                    if let Some(enabled) = mode {
                        self.nocaseglob_enabled = enabled;
                    }
                }
                "nocasematch" => {
                    if let Some(enabled) = mode {
                        self.nocasematch_enabled = enabled;
                    }
                }
                "patsub_replacement" => {
                    if let Some(enabled) = mode {
                        self.patsub_replacement_enabled = enabled;
                    }
                }
                "lastpipe" => {
                    if let Some(enabled) = mode {
                        self.lastpipe_enabled = enabled;
                    }
                }
                "extglob" => {}
                _ => status = 1,
            }
        }
        Ok(Lowered::value(Cmd::status(int(status))))
    }

    fn walk_echo(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut newline = true;
        let mut escapes = false;
        let mut rest = Vec::new();
        let mut in_opts = true;
        for w in suffix {
            if in_opts {
                if let Some(t) = literal_word_text(&w) {
                    if t.len() > 1
                        && t.starts_with('-')
                        && t[1..].chars().all(|c| matches!(c, 'n' | 'e' | 'E'))
                    {
                        if t.contains('n') {
                            newline = false;
                        }
                        if t.contains('e') {
                            escapes = true;
                        }
                        if t.contains('E') {
                            escapes = false;
                        }
                        continue;
                    }
                }
                in_opts = false;
            }
            rest.push(w);
        }
        let mut args = self.words_shell_args(rest)?;
        if escapes {
            for a in &mut args {
                if let ExprKind::Lit(Literal::Str(s)) = &a.value.kind {
                    a.value = lit(&decode_ansi_c(s));
                }
            }
        }
        if args.iter().all(|arg| !arg.spread) {
            let mut fields = Vec::new();
            let mut all_literal = true;
            for arg in &args {
                if let Some(text) = literal_string(&arg.value) {
                    fields.push(text.to_string());
                } else {
                    all_literal = false;
                    break;
                }
            }
            if all_literal {
                let mut text = fields.join(" ");
                if newline {
                    text.push('\n');
                }
                return Ok(Lowered::value(Cmd::status(bash_stdout_status_expr(lit(
                    &text,
                )))));
            }
        }
        let mut plain_args: Vec<Expression> = args.iter().map(|arg| arg.value.clone()).collect();
        let text = if newline {
            parts_to_expr(vec![
                Part::Expr(array_join(shell_args_array(args), lit(" "))),
                Part::Text("\n".to_string()),
            ])
        } else {
            let joined = match plain_args.len() {
                0 => lit(""),
                1 if !args.iter().any(|arg| arg.spread) => plain_args.remove(0),
                _ => array_join(shell_args_array(args), lit(" ")),
            };
            joined
        };
        Ok(Lowered::value(Cmd::status(bash_stdout_status_expr(text))))
    }

    fn walk_mkdir(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut recursive = false;
        let mut paths = Vec::new();
        for w in suffix {
            if let Some(text) = literal_word_text(&w) {
                if text.starts_with('-') {
                    if text.contains('p') {
                        recursive = true;
                    }
                    continue;
                }
            }
            paths.extend(self.words_exprs(vec![w])?);
        }
        if paths.is_empty() {
            return Ok(Lowered::value(Cmd::status(int(1))));
        }
        let mkdir_name = if recursive {
            "__bash_mkdir_all"
        } else {
            "__bash_mkdir"
        };
        let calls = paths
            .into_iter()
            .map(|path| call_named(mkdir_name, vec![bash_path_expr(path)]))
            .collect();
        Ok(Lowered::value(Cmd::boolean(sequence(calls))))
    }

    /// `printf [-v var] format args…` — backslash escapes in a literal format
    /// are decoded here; `%` conversions are the shared formatter's.
    fn walk_printf(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut it = suffix.into_iter();
        let mut target: Option<String> = None;
        let mut first = it.next();
        if let Some(w) = &first {
            if literal_word_text(w).as_deref() == Some("-v") {
                target = it.next().and_then(|w| self.static_word_text(&w));
                first = it.next();
            }
        }
        let Some(fmt_word) = first else {
            return Ok(Lowered::value(Cmd::status(int(1))));
        };
        let mut fmt = self.word_expr(fmt_word)?;
        let mut fmt_text = None;
        if let ExprKind::Lit(Literal::Str(s)) = &fmt.kind {
            fmt_text = Some(decode_printf_escapes(s));
        }
        if let Some(s) = &fmt_text {
            fmt = lit(s);
        }
        let shell_args = self.words_shell_args(it.collect())?;
        let sprintf = self.bash_printf_expr(fmt, fmt_text.as_deref(), shell_args);
        Ok(match target {
            Some(name) => {
                if !is_unset_target(&name)
                    || is_readonly_parameter_target(&name)
                    || self.target_text_is_readonly(&name)
                {
                    return Ok(Lowered::stmts(vec![assign_stmt(
                        ident("__bash_status"),
                        int(1),
                    )]));
                }
                let t = self.name_or_element_target(&name)?;
                Lowered::stmts(vec![
                    assign_stmt(t, sprintf),
                    assign_stmt(ident("__bash_status"), int(0)),
                ])
            }
            None => Lowered::value(Cmd::status(bash_stdout_status_expr(sprintf))),
        })
    }

    fn printf_value(&mut self, suffix: Vec<Pair<Rule>>) -> R<Option<Expression>> {
        let mut it = suffix.into_iter();
        if let Some(w) = it.next() {
            if literal_word_text(&w).as_deref() == Some("-v") {
                return Ok(None);
            }
            let mut fmt = self.word_expr(w)?;
            let mut fmt_text = None;
            if let ExprKind::Lit(Literal::Str(s)) = &fmt.kind {
                fmt_text = Some(decode_printf_escapes(s));
            }
            if let Some(s) = &fmt_text {
                fmt = lit(s);
            }
            let shell_args = self.words_shell_args(it.collect())?;
            return Ok(Some(self.bash_printf_expr(fmt, fmt_text.as_deref(), shell_args)));
        }
        Ok(Some(lit("")))
    }

    fn bash_printf_expr(
        &mut self,
        fmt: Expression,
        fmt_text: Option<&str>,
        shell_args: Vec<ShellArg>,
    ) -> Expression {
        let Some(fmt_text) = fmt_text else {
            return bash_sprintf_array(fmt, shell_args);
        };
        let conversions = printf_conversions(fmt_text);
        let common_fmt = if conversions.iter().any(|conv| *conv == 'b') {
            lit(&printf_common_format(fmt_text))
        } else {
            fmt.clone()
        };
        if conversions.len() == 1 {
            let conv = conversions[0];
            if shell_args.len() == 1 && shell_args[0].spread {
                let item = self.fresh("__bash_printf_arg");
                let rendered = bash_sprintf_array(
                    common_fmt.clone(),
                    vec![ShellArg {
                        value: convert_bash_printf_arg(conv, ident(&item)),
                        spread: false,
                    }],
                );
                let rendered_items = method(
                    shell_args[0].value.clone(),
                    "map",
                    vec![lambda_expr(vec![param_named(&item)], rendered)],
                );
                return ternary(
                    binary(
                        BinOp::StrictEq,
                        member(shell_args[0].value.clone(), "length"),
                        int(0),
                    ),
                    bash_sprintf_array(
                        common_fmt,
                        vec![ShellArg {
                            value: printf_default_for_conversion(conv),
                            spread: false,
                        }],
                    ),
                    array_join(rendered_items, lit("")),
                );
            }
            if !shell_args.iter().any(|arg| arg.spread) {
                let args = if shell_args.is_empty() {
                    vec![ShellArg {
                        value: printf_default_for_conversion(conv),
                        spread: false,
                    }]
                } else {
                    shell_args
                };
                return parts_to_expr(
                    args.into_iter()
                        .map(|arg| {
                            Part::Expr(bash_sprintf_array(
                                common_fmt.clone(),
                                vec![ShellArg {
                                    value: convert_bash_printf_arg(conv, arg.value),
                                    spread: false,
                                }],
                            ))
                        })
                        .collect(),
                );
            }
        }
        let has_spread = shell_args.iter().any(|arg| arg.spread);
        let provided = shell_args.len();
        let mut args = shell_args;
        if !has_spread {
            args.extend(printf_missing_default_args(fmt_text, provided).into_iter().map(
                |value| ShellArg {
                    value,
                    spread: false,
                },
            ));
        }
        bash_sprintf_array(common_fmt, args)
    }

    /// `read [-r] [-a name] [-d c] [-p prompt] [-t n] [-s] [-n k] [-u fd] names…`
    ///
    /// Consumes one line of `__bash_stdin`; true when input was available.
    fn walk_read(&mut self, suffix: Vec<Pair<Rule>>, ifs_empty: bool) -> R<Lowered> {
        let mut array_name: Option<String> = None;
        let mut input_fd: Option<String> = None;
        let mut read_until_eof = false;
        let mut raw = false;
        let mut names: Vec<String> = Vec::new();
        let mut it = suffix.into_iter();
        while let Some(w) = it.next() {
            let Some(t) = self.static_word_text(&w) else {
                continue;
            };
            if t.starts_with('-') && t.len() > 1 {
                let flags: Vec<char> = t[1..].chars().collect();
                if flags.contains(&'r') {
                    raw = true;
                }
                // Flags that take an argument consume the next word.
                if let Some(f) = flags.last() {
                    match f {
                        'a' => array_name = it.next().and_then(|w| self.static_word_text(&w)),
                        'd' => {
                            let delimiter = it
                                .next()
                                .and_then(|w| self.static_word_text(&w))
                                .unwrap_or_default();
                            if delimiter.is_empty() {
                                read_until_eof = true;
                            }
                        }
                        'u' => {
                            input_fd = it.next().and_then(|w| self.static_word_text(&w));
                        }
                        'p' | 't' | 'n' | 'N' | 'i' => {
                            it.next();
                        }
                        _ => {}
                    }
                }
                continue;
            }
            names.push(t);
        }
        if names.is_empty() && array_name.is_none() {
            names.push("REPLY".into());
        }
        for name in &names {
            self.record_variable(name);
        }
        if let Some(name) = &array_name {
            self.record_variable(name);
        }
        let line = ident("__bash_ret");
        let input_target = input_fd
            .as_deref()
            .map(|fd| index(ident("__bash_fds"), lit(fd)))
            .unwrap_or_else(|| ident("__bash_stdin"));
        let input_value = binary(BinOp::NullCoalesce, input_target.clone(), lit(""));
        for name in names.iter().chain(array_name.iter()) {
            if is_readonly_parameter_target(name) || self.target_text_is_readonly(name) {
                return Ok(Lowered::stmts(vec![assign_stmt(
                    ident("__bash_status"),
                    int(1),
                )]));
            }
        }
        let mut seq = if read_until_eof {
            vec![
                assign_expr(ident("__bash_line"), int(-1)),
                assign_expr(line.clone(), input_value.clone()),
            ]
        } else {
            vec![
                assign_expr(
                    ident("__bash_line"),
                    method(input_value.clone(), "indexOf", vec![lit("\n")]),
                ),
                assign_expr(
                    line.clone(),
                    ternary(
                        binary(BinOp::Lt, ident("__bash_line"), int(0)),
                        input_value.clone(),
                        method(
                            input_value.clone(),
                            "slice",
                            vec![int(0), ident("__bash_line")],
                        ),
                    ),
                ),
            ]
        };
        let had = binary(BinOp::StrictNotEq, input_value.clone(), lit(""));
        let had_var = "__bash_had".to_string();
        seq.push(assign_expr(ident(&had_var), had));
        seq.push(assign_expr(
            input_target,
            if read_until_eof {
                lit("")
            } else {
                ternary(
                    binary(BinOp::Lt, ident("__bash_line"), int(0)),
                    lit(""),
                    method(
                        input_value,
                        "slice",
                        vec![binary(BinOp::Add, ident("__bash_line"), int(1))],
                    ),
                )
            },
        ));
        if !raw {
            seq.push(assign_expr(
                line.clone(),
                regex_replace_expr(line.clone(), "\\\\(.)", "g", lit("$1")),
            ));
        }
        let trimmed = if ifs_empty {
            line.clone()
        } else {
            method(line.clone(), "trim", vec![])
        };
        let fields = if ifs_empty {
            array(vec![trimmed.clone()])
        } else {
            ternary(
                binary(
                    BinOp::Or,
                    binary(BinOp::StrictEq, ident("IFS"), undefined()),
                    binary(BinOp::StrictEq, ident("IFS"), lit(" \t\n")),
                ),
                call_named("__bash_re_split", vec![trimmed.clone(), lit("[ \t]+")]),
                method(
                    trimmed.clone(),
                    "split",
                    vec![method(param_value(ident("IFS")), "charAt", vec![int(0)])],
                ),
            )
        };
        if let Some(arr) = array_name {
            let t = self.name_or_element_target(&arr)?;
            seq.push(assign_expr(
                t,
                ternary(
                    binary(BinOp::StrictEq, trimmed.clone(), lit("")),
                    array(Vec::new()),
                    fields.clone(),
                ),
            ));
        } else if names.len() == 1 {
            let t = self.name_or_element_target(&names[0])?;
            seq.push(assign_expr(t, trimmed.clone()));
        } else {
            let fields_var = "__bash_fields".to_string();
            seq.push(assign_expr(ident(&fields_var), fields));
            let n = names.len();
            for (i, name) in names.iter().enumerate() {
                let t = self.name_or_element_target(name)?;
                let value = if i + 1 == n {
                    array_join(
                        method(ident(&fields_var), "slice", vec![int(i as i64)]),
                        lit(" "),
                    )
                } else {
                    binary(
                        BinOp::NullCoalesce,
                        index(ident(&fields_var), int(i as i64)),
                        lit(""),
                    )
                };
                seq.push(assign_expr(t, value));
            }
        }
        seq.push(ident(&had_var));
        // The temporaries are module-level variables; each `read` overwrites them.
        Ok(Lowered::value(Cmd::boolean(sequence(seq))))
    }

    /// `mapfile [-t] [-d c] [-n k] [-O k] [-s k] name` — the lines of `__bash_stdin`.
    fn walk_mapfile(&mut self, suffix: Vec<Pair<Rule>>) -> R<Lowered> {
        let mut name = "MAPFILE".to_string();
        let mut it = suffix.into_iter();
        while let Some(w) = it.next() {
            let Some(t) = literal_word_text(&w) else {
                continue;
            };
            if t.starts_with('-') && t.len() > 1 {
                if matches!(
                    t.chars().last(),
                    Some('d' | 'n' | 'O' | 's' | 'u' | 'C' | 'c')
                ) {
                    it.next();
                }
                continue;
            }
            name = t;
        }
        let t = self.name_or_element_target(&name)?;
        let lines = method(
            method(
                ident("__bash_stdin"),
                "replace",
                vec![regexp(lit("\\n$"), ""), lit("")],
            ),
            "split",
            vec![lit("\n")],
        );
        Ok(Lowered::value(Cmd::boolean(sequence(vec![
            assign_expr(
                t,
                ternary(
                    binary(BinOp::StrictEq, ident("__bash_stdin"), lit("")),
                    array(Vec::new()),
                    lines,
                ),
            ),
            assign_expr(ident("__bash_stdin"), lit("")),
            Expression::bool(true),
        ]))))
    }

    /// `local`/`declare`/`typeset`/`readonly`/`export [-flags] name[=value]…`.
    fn walk_declaration(&mut self, which: &str, suffix: Vec<Pair<Rule>>) -> R<Vec<Statement>> {
        if which == "local" && self.function_depth == 0 {
            return Ok(vec![assign_stmt(ident("__bash_status"), int(1))]);
        }
        let mut flags = String::new();
        let mut items = Vec::new();
        for w in suffix {
            if w.as_rule() == Rule::word {
                if let Some(t) = self.static_word_text(&w) {
                    if t.starts_with('-') || t.starts_with('+') {
                        flags.push_str(&t);
                        continue;
                    }
                }
            }
            items.push(w);
        }
        let kind = if self.function_depth > 0 {
            VarDeclKind::FunctionScoped
        } else {
            VarDeclKind::Var
        };
        let default_init = if flag_enabled(&flags, 'A') {
            Some(object(Vec::new()))
        } else if flag_enabled(&flags, 'a') {
            Some(array(Vec::new()))
        } else {
            None
        };
        let global_binding =
            self.function_depth > 0 && (which == "readonly" || flag_enabled(&flags, 'g'));
        // `declare -p` prints tracked declaration metadata; no runtime helper.
        if flags.contains('p') {
            if which == "export" {
                return Ok(self.walk_export_print());
            }
            if which == "readonly" {
                return Ok(self.walk_readonly_print());
            }
            return self.walk_declare_print(&items);
        }
        if which == "export" && flag_enabled(&flags, 'f') {
            let mut status = 0;
            for item in items {
                let Some(name) = literal_word_text(&item) else {
                    status = 1;
                    continue;
                };
                if self.functions.contains_key(&name) {
                    self.exported_functions.insert(name);
                } else {
                    status = 1;
                }
            }
            return Ok(vec![assign_stmt(ident("__bash_status"), int(status))]);
        }
        if (which == "declare" || which == "typeset")
            && flag_enabled(&flags, 't')
            && flag_enabled(&flags, 'f')
        {
            for item in items {
                if let Some(name) = literal_word_text(&item) {
                    if self.functions.contains_key(&name) {
                        self.traced_functions.insert(name);
                    }
                }
            }
            return Ok(vec![assign_stmt(ident("__bash_status"), int(0))]);
        }
        // `declare -f` / `-F` only print; minimal status for now.
        if flag_enabled(&flags, 'F') {
            return Ok(self.walk_function_print(&items, false));
        }
        if flag_enabled(&flags, 'f') {
            return Ok(self.walk_function_print(&items, true));
        }
        let mut declarations = Vec::new();
        let mut extra = Vec::new();
        let mut global_names = Vec::new();
        let mut status = 0;
        for item in items {
            match item.as_rule() {
                Rule::assignment_word => {
                    let (target, op, value) = self
                        .assignment_parts_with_array_kind(item.clone(), flag_enabled(&flags, 'A'))?;
                    match target {
                        AssignTarget::Name(n) => {
                            let mut value = value;
                            if is_readonly_parameter_target(&n) {
                                status = 1;
                                continue;
                            }
                            if self.target_text_is_readonly(&n) {
                                status = 1;
                                continue;
                            }
                            if global_binding {
                                global_names.push(n.clone());
                            }
                            self.record_variable(&n);
                            if which == "readonly" {
                                self.readonly_vars.insert(n.clone());
                            }
                            if flag_enabled(&flags, 'A') {
                                self.assoc_arrays.insert(n.clone());
                                self.indexed_arrays.remove(&n);
                            } else if flag_enabled(&flags, 'a') {
                                self.indexed_arrays.insert(n.clone());
                                self.assoc_arrays.remove(&n);
                            } else if which == "local" {
                                self.assoc_arrays.remove(&n);
                                self.indexed_arrays.remove(&n);
                            }
                            if flag_enabled(&flags, 'i') {
                                self.integer_vars.insert(n.clone());
                            } else if which == "local" {
                                self.integer_vars.remove(&n);
                            }
                            if flag_enabled(&flags, 'l') {
                                self.lowercase_vars.insert(n.clone());
                                self.uppercase_vars.remove(&n);
                            } else if which == "local" {
                                self.lowercase_vars.remove(&n);
                            }
                            if flag_enabled(&flags, 'u') {
                                self.uppercase_vars.insert(n.clone());
                                self.lowercase_vars.remove(&n);
                            } else if which == "local" {
                                self.uppercase_vars.remove(&n);
                            }
                            if flag_disabled(&flags, 'i') {
                                self.integer_vars.remove(&n);
                            }
                            if flag_disabled(&flags, 'l') {
                                self.lowercase_vars.remove(&n);
                            }
                            if flag_disabled(&flags, 'u') {
                                self.uppercase_vars.remove(&n);
                            }
                            if which == "export"
                                || flag_enabled(&flags, 'x')
                                || flag_disabled(&flags, 'x')
                            {
                                if (flag_enabled(&flags, 'n') && which == "export")
                                    || flag_disabled(&flags, 'x')
                                {
                                    self.exported_vars.remove(&n);
                                } else {
                                    self.exported_vars.insert(n.clone());
                                }
                            }
                            if flag_enabled(&flags, 'n') {
                                let nameref_target = literal_string(&value)
                                    .map(str::to_string)
                                    .or_else(|| self.assignment_word_static_rhs_text(&item));
                                if let Some(target_name) = nameref_target {
                                    if self.nameref_would_cycle(&n, &target_name) {
                                        extra.push(expr_stmt(call_named(
                                            "printf",
                                            vec![
                                                lit("%s\n"),
                                                lit("bash: warning: circular name reference"),
                                            ],
                                        )));
                                    }
                                    self.namerefs.insert(n.clone(), target_name.clone());
                                    self.variable_values.insert(n.clone(), target_name.clone());
                                    value = lit(&target_name);
                                }
                            } else {
                                if which == "local" {
                                    self.namerefs.remove(&n);
                                }
                                if op != "+=" {
                                    if let Some(snapshot) =
                                        bash_array_snapshot(&value, self.assoc_arrays.contains(&n))
                                    {
                                        self.array_values.insert(n.clone(), snapshot);
                                        self.variable_values.remove(&n);
                                    } else if let Some(s) = literal_string(&value) {
                                        self.variable_values.insert(n.clone(), s.to_string());
                                        self.array_values.remove(&n);
                                    } else {
                                        self.variable_values.remove(&n);
                                        self.array_values.remove(&n);
                                    }
                                }
                            }
                            if op == "+=" {
                                extra.push(if self.integer_vars.contains(&n) {
                                    assign_stmt(
                                        ident(&n),
                                        call_named(
                                            "__bash_string",
                                            vec![binary(
                                                BinOp::Add,
                                                to_number(ident(&n)),
                                                to_number(value),
                                            )],
                                        ),
                                    )
                                } else {
                                    compound_append(ident(&n), value)
                                });
                            } else if which == "export" {
                                let value = self.apply_variable_attributes(&n, value)?;
                                extra.push(assign_stmt(ident(&n), value));
                            } else if global_binding {
                                let value = self.apply_variable_attributes(&n, value)?;
                                extra.push(assign_stmt(ident(&n), value));
                            } else {
                                let value = self.apply_variable_attributes(&n, value)?;
                                declarations.push(declarator(&n, Some(value)));
                            }
                        }
                        AssignTarget::Element(arr, key) => {
                            if is_readonly_parameter_target(&arr)
                                || self.target_text_is_readonly(&arr)
                            {
                                status = 1;
                                continue;
                            }
                            extra.push(assign_stmt(index(ident(&arr), key), value));
                        }
                    }
                }
                _ => {
                    let raw_item = item.as_str().trim();
                    if let Some((raw_name, _)) = raw_item.split_once('=') {
                        let raw_name = raw_name.strip_suffix('+').unwrap_or(raw_name);
                        if is_name(raw_name) {
                            if let Ok(mut parsed) =
                                WashmParser::parse(Rule::assignment_word, raw_item)
                            {
                                if let Some(parsed) = parsed.next() {
                                    let lowered = self.walk_declaration(which, vec![parsed])?;
                                    extra.extend(lowered);
                                    continue;
                                }
                            }
                        }
                    }
                    if let Some(n) = self.static_word_text(&item) {
                        if let Some((name, value_text)) = n.split_once('=') {
                            if is_name(name) {
                                let synthetic = format!("{name}={value_text}");
                                let parsed = WashmParser::parse(Rule::assignment_word, &synthetic)
                                    .map_err(|e| format!("dynamic declaration assignment: {e}"))?
                                    .next()
                                    .ok_or("dynamic declaration assignment missing")?;
                                let lowered = self.walk_declaration(which, vec![parsed])?;
                                extra.extend(lowered);
                                continue;
                            }
                            status = 1;
                            continue;
                        }
                        if is_name(&n) {
                            let existed = self.variables.contains(&n)
                                || self.variable_values.contains_key(&n)
                                || self.array_values.contains_key(&n);
                            if global_binding {
                                global_names.push(n.clone());
                            }
                            self.record_variable(&n);
                            if which == "readonly" {
                                self.readonly_vars.insert(n.clone());
                            }
                            if which == "export" {
                                if flag_enabled(&flags, 'n') {
                                    self.exported_vars.remove(&n);
                                } else {
                                    self.exported_vars.insert(n);
                                }
                                continue;
                            }
                            if flag_enabled(&flags, 'A') {
                                self.assoc_arrays.insert(n.clone());
                                self.indexed_arrays.remove(&n);
                            }
                            if flag_enabled(&flags, 'a') {
                                self.indexed_arrays.insert(n.clone());
                                self.assoc_arrays.remove(&n);
                            }
                            if which == "local"
                                && !flag_enabled(&flags, 'A')
                                && !flag_enabled(&flags, 'a')
                            {
                                self.assoc_arrays.remove(&n);
                                self.indexed_arrays.remove(&n);
                            }
                            if flag_enabled(&flags, 'i') {
                                self.integer_vars.insert(n.clone());
                            } else if which == "local" {
                                self.integer_vars.remove(&n);
                            }
                            if which == "local" && !flag_enabled(&flags, 'n') {
                                self.namerefs.remove(&n);
                            }
                            if flag_enabled(&flags, 'l') {
                                self.lowercase_vars.insert(n.clone());
                                self.uppercase_vars.remove(&n);
                            } else if which == "local" {
                                self.lowercase_vars.remove(&n);
                            }
                            if flag_enabled(&flags, 'u') {
                                self.uppercase_vars.insert(n.clone());
                                self.lowercase_vars.remove(&n);
                            } else if which == "local" {
                                self.uppercase_vars.remove(&n);
                            }
                            if flag_disabled(&flags, 'r') && self.readonly_vars.contains(&n) {
                                status = 1;
                                continue;
                            }
                            if flag_disabled(&flags, 'i') {
                                self.integer_vars.remove(&n);
                            }
                            if flag_disabled(&flags, 'l') {
                                self.lowercase_vars.remove(&n);
                            }
                            if flag_disabled(&flags, 'u') {
                                self.uppercase_vars.remove(&n);
                            }
                            if default_init.is_some() || !flags.contains('+') {
                                if let Some(init) = default_init.clone() {
                                    if let Some(snapshot) =
                                        bash_array_snapshot(&init, self.assoc_arrays.contains(&n))
                                    {
                                        self.array_values.insert(n.clone(), snapshot);
                                    }
                                    if global_binding {
                                        extra.push(assign_stmt(ident(&n), init));
                                    } else {
                                        declarations.push(declarator(&n, Some(init)));
                                    }
                                } else if which == "local" || !existed {
                                    if global_binding {
                                        extra.push(assign_stmt(ident(&n), undefined()));
                                    } else {
                                        declarations.push(declarator(&n, Some(undefined())));
                                    }
                                    self.variable_values.remove(&n);
                                }
                            }
                            if which == "export"
                                || flag_enabled(&flags, 'x')
                                || flag_disabled(&flags, 'x')
                            {
                                if (flag_enabled(&flags, 'n') && which == "export")
                                    || flag_disabled(&flags, 'x')
                                {
                                    self.exported_vars.remove(&n);
                                } else {
                                    self.exported_vars.insert(n);
                                }
                            }
                        } else {
                            status = 1;
                        }
                    } else {
                        status = 1;
                    }
                }
            }
        }
        let mut out = Vec::new();
        if !global_names.is_empty() {
            global_names.sort();
            global_names.dedup();
            out.push(Statement::new(StmtKind::ScopeDecl {
                kind: ScopeDeclKind::Global,
                names: global_names,
            }));
        }
        if !declarations.is_empty() && self.function_depth == 0 {
            for decl in declarations {
                if let BindingPattern::Ident(name) = decl.pattern {
                    extra.insert(0, assign_stmt(ident(&name), decl.init.unwrap_or_else(undefined)));
                }
            }
        } else if !declarations.is_empty() {
            out.push(Statement::new(StmtKind::VarDecl { declarations, kind }));
        }
        out.extend(extra);
        if out.is_empty() {
            out.push(assign_stmt(ident("__bash_status"), int(status)));
        } else {
            out.push(assign_stmt(ident("__bash_status"), int(status)));
        }
        Ok(out)
    }

    fn walk_declare_print(&mut self, items: &[Pair<Rule>]) -> R<Vec<Statement>> {
        let mut exprs = Vec::new();
        let mut status = 0;
        for item in items {
            let Some(name) = literal_word_text(item) else {
                status = 1;
                continue;
            };
            if let Some(target) = self.namerefs.get(&name) {
                exprs.push(call_named(
                    "printf",
                    vec![lit("declare -n %s=\"%s\"\n"), lit(&name), lit(target)],
                ));
            } else if self.assoc_arrays.contains(&name) {
                if let Some(entries) = self.array_values.get(&name) {
                    let payload = bash_array_declare_payload(entries, true);
                    exprs.push(call_named(
                        "printf",
                        vec![lit("declare -A %s=(%s)\n"), lit(&name), lit(&payload)],
                    ));
                } else {
                    exprs.push(call_named(
                        "printf",
                        vec![lit("declare -A %s=()\n"), lit(&name)],
                    ));
                }
            } else if self.indexed_arrays.contains(&name) {
                if let Some(entries) = self.array_values.get(&name) {
                    let payload = bash_array_declare_payload(entries, true);
                    exprs.push(call_named(
                        "printf",
                        vec![lit("declare -a %s=(%s)\n"), lit(&name), lit(&payload)],
                    ));
                } else {
                    exprs.push(call_named(
                        "printf",
                        vec![lit("declare -a %s=()\n"), lit(&name)],
                    ));
                }
            } else if self.variables.contains(&name) || self.variable_values.contains_key(&name) {
                let attrs = self.attribute_letters(&name);
                let flag = if attrs.is_empty() {
                    "--".to_string()
                } else {
                    format!("-{attrs}")
                };
                exprs.push(call_named(
                    "printf",
                    vec![
                        lit("declare %s %s=\"%s\"\n"),
                        lit(&flag),
                        lit(&name),
                        param_value(ident(&name)),
                    ],
                ));
            } else {
                status = 1;
            }
        }
        exprs.push(int(status));
        Ok(vec![
            expr_stmt(sequence(exprs)),
            assign_stmt(ident("__bash_status"), int(status)),
        ])
    }

    fn walk_readonly_print(&mut self) -> Vec<Statement> {
        let mut exprs = Vec::new();
        let mut names: Vec<String> = self.readonly_vars.iter().cloned().collect();
        names.sort();
        for name in names {
            if self.assoc_arrays.contains(&name) {
                let payload = self
                    .array_values
                    .get(&name)
                    .map(|entries| bash_array_declare_payload(entries, true))
                    .unwrap_or_default();
                exprs.push(call_named(
                    "printf",
                    vec![lit("declare -rA %s=(%s)\n"), lit(&name), lit(&payload)],
                ));
            } else if self.indexed_arrays.contains(&name) {
                let payload = self
                    .array_values
                    .get(&name)
                    .map(|entries| bash_array_declare_payload(entries, true))
                    .unwrap_or_default();
                exprs.push(call_named(
                    "printf",
                    vec![lit("declare -ra %s=(%s)\n"), lit(&name), lit(&payload)],
                ));
            } else {
                let value = self.variable_values.get(&name).cloned().unwrap_or_default();
                exprs.push(call_named(
                    "printf",
                    vec![lit("declare -r %s=\"%s\"\n"), lit(&name), lit(&value)],
                ));
            }
        }
        exprs.push(int(0));
        vec![
            expr_stmt(sequence(exprs)),
            assign_stmt(ident("__bash_status"), int(0)),
        ]
    }

    fn walk_function_print(&mut self, items: &[Pair<Rule>], body: bool) -> Vec<Statement> {
        let mut exprs = Vec::new();
        let mut status = 0;
        let mut names: Vec<String> = if items.is_empty() {
            self.functions.keys().cloned().collect()
        } else {
            items.iter().filter_map(literal_word_text).collect()
        };
        names.sort();
        for name in names {
            if self.functions.contains_key(&name) {
                if body {
                    let display = self
                        .function_bodies
                        .get(&name)
                        .cloned()
                        .unwrap_or_else(|| format!("{name} () {{ :; }}"));
                    exprs.push(bash_stdout_write_expr(parts_to_expr(vec![
                        Part::Text(display),
                        Part::Text("\n".into()),
                    ])));
                } else if self.traced_functions.contains(&name) {
                    exprs.push(bash_stdout_write_expr(lit(&format!(
                        "declare -ft {name}\n"
                    ))));
                } else {
                    exprs.push(bash_stdout_write_expr(lit(&format!("{name}\n"))));
                }
            } else if self.variables.contains(&name) {
                status = 1;
            } else {
                status = 1;
            }
        }
        exprs.push(int(status));
        vec![
            expr_stmt(sequence(exprs)),
            assign_stmt(ident("__bash_status"), int(status)),
        ]
    }

    fn walk_export_print(&mut self) -> Vec<Statement> {
        let mut exprs = Vec::new();
        let mut names: Vec<String> = self.exported_vars.iter().cloned().collect();
        names.sort();
        for name in names {
            let value = self.variable_values.get(&name).cloned().unwrap_or_default();
            exprs.push(call_named(
                "printf",
                vec![lit("declare -x %s=\"%s\"\n"), lit(&name), lit(&value)],
            ));
        }
        exprs.push(int(0));
        vec![
            expr_stmt(sequence(exprs)),
            assign_stmt(ident("__bash_status"), int(0)),
        ]
    }

    fn attribute_letters(&self, name: &str) -> String {
        let resolved = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        let mut attrs = String::new();
        if self.assoc_arrays.contains(&resolved) {
            attrs.push('A');
        } else if self.indexed_arrays.contains(&resolved) {
            attrs.push('a');
        }
        if self.readonly_vars.contains(&resolved) {
            attrs.push('r');
        }
        if self.integer_vars.contains(&resolved) {
            attrs.push('i');
        }
        if self.lowercase_vars.contains(&resolved) {
            attrs.push('l');
        }
        if self.uppercase_vars.contains(&resolved) {
            attrs.push('u');
        }
        if self.exported_vars.contains(&resolved) {
            attrs.push('x');
        }
        attrs
    }

    fn name_or_element_target(&mut self, text: &str) -> R<Expression> {
        if let Some((arr, rest)) = text.split_once('[') {
            let key = rest.trim_end_matches(']');
            let target = self
                .resolve_nameref_text(arr)
                .unwrap_or_else(|| arr.to_string());
            if target.contains('[') {
                return self.name_or_element_target(&target);
            }
            let key = if self.assoc_arrays.contains(&target) {
                self.subscript_assoc_key_expr(key)?
            } else {
                arith_index(self.subscript_text_expr(key)?)
            };
            return Ok(index(ident(&target), key));
        }
        if let Some(target) = self.resolve_nameref_text(text) {
            return self.name_or_element_target(&target);
        }
        Ok(ident(text))
    }

    fn resolve_nameref_text(&self, name: &str) -> Option<String> {
        let mut current = name.to_string();
        let mut seen = HashSet::new();
        let mut resolved = None;
        while seen.insert(current.clone()) {
            let Some(next) = self.namerefs.get(&current) else {
                break;
            };
            if is_name(next) && seen.contains(next) {
                return None;
            }
            resolved = Some(next.clone());
            if !is_name(next) {
                break;
            }
            current = next.clone();
        }
        resolved
    }

    // ── redirections ─────────────────────────────────────────────────────

    /// Apply redirections around `stmts`.
    ///
    /// stdin forms set `__bash_stdin` for the duration of the command; stdout
    /// forms capture the output and discard it (`/dev/null`) or write it to
    /// the target file. fd-2 forms are normalized to shell-state flags that
    /// the in-process output builtins consult before writing to the funnel.
    fn apply_redirections(
        &mut self,
        stmts: Vec<Statement>,
        redirs: &[Pair<Rule>],
    ) -> R<Vec<Statement>> {
        let mut stdin: Option<Expression> = None;
        let mut stdout: Option<(bool, Expression)> = None; // (append, target)
        let mut stdout_procsub: Option<Expression> = None;
        let mut stdout_fd: Option<Expression> = None;
        let mut stdout_null = false;
        let mut stdout_to_stderr = false;
        let mut stderr_to_stdout: Option<bool> = None;
        let mut stderr_null: Option<bool> = None;
        let mut prelude = Vec::new();
        for r in redirs {
            let (fd, op, target, mut before) = self.redirection_parts(r.clone())?;
            prelude.append(&mut before);
            match op.as_str() {
                "<<<" => {
                    stdin = Some(parts_to_expr(vec![
                        Part::Expr(target),
                        Part::Text("\n".into()),
                    ]))
                }
                "<<" => stdin = Some(target),
                "<" if fd == "0" => {
                    stdin = Some(match target.kind {
                        ExprKind::Call { .. } if is_procsub(&target) => procsub_capture(target),
                        _ => self.read_file_or_fd_expr(target),
                    })
                }
                "<&" if fd == "0" && literal_string(&target) != Some("-") => {
                    stdin = Some(self.read_fd_expr(target));
                }
                ">" | ">|" | ">>" | "&>" | "&>>" if fd == "1" || op.starts_with('&') => {
                    let append = op.ends_with(">>");
                    match &target.kind {
                        ExprKind::Lit(Literal::Str(s)) if s == "/dev/null" => {
                            stdout_to_stderr = false;
                            stdout_null = true;
                            stdout = None;
                            if op.starts_with('&') {
                                stderr_null = Some(true);
                                stderr_to_stdout = Some(false);
                            }
                        }
                        ExprKind::Lit(Literal::Str(s))
                            if s == "/dev/stdout" || s == "/dev/stderr" =>
                        {
                            stdout_to_stderr = false;
                        }
                        ExprKind::Call { .. } if is_procsub(&target) => {
                            stdout_to_stderr = false;
                            stdout_null = false;
                            stdout = None;
                            stdout_procsub = Some(target);
                            if op.starts_with('&') {
                                stderr_null = Some(false);
                                stderr_to_stdout = Some(true);
                            }
                        }
                        _ => {
                            stdout_to_stderr = false;
                            stdout_null = false;
                            stdout = Some((append, target));
                            if op.starts_with('&') {
                                stderr_null = Some(false);
                                stderr_to_stdout = Some(true);
                            }
                        }
                    }
                }
                ">" | ">|" | ">>" if fd == "2" => {
                    if matches!(&target.kind, ExprKind::Lit(Literal::Str(s)) if s == "/dev/null") {
                        stderr_null = Some(true);
                        stderr_to_stdout = Some(false);
                    }
                }
                ">&" if fd == "1" && literal_string(&target) == Some("2") => {
                    stdout_to_stderr = true;
                }
                ">&" if fd == "1" => {
                    stdout_to_stderr = false;
                    stdout_null = false;
                    stdout_fd = Some(target);
                }
                ">&" if fd == "2" && literal_string(&target) == Some("1") => {
                    if stdout_null {
                        stderr_null = Some(true);
                        stderr_to_stdout = Some(false);
                    } else {
                        stderr_null = Some(false);
                        stderr_to_stdout = Some(true);
                    }
                }
                _ => {}
            }
        }
        let mut out = prelude;
        out.extend(stmts);
        if stdout_null {
            let saved = self.fresh("__bash_saved_stdout_null");
            let mut wrapped = vec![
                let_stmt(&saved, ident("__bash_stdout_null")),
                assign_stmt(ident("__bash_stdout_null"), Expression::bool(true)),
            ];
            wrapped.extend(out);
            wrapped.push(assign_stmt(ident("__bash_stdout_null"), ident(&saved)));
            out = wrapped;
        } else if let Some((append, target)) = stdout {
            let target_var = self.fresh("__bash_redir_target");
            let content = self.fresh("__bash_redir_content");
            let fd_key = self.fresh("__bash_redir_fd_key");
            let coproc = self.fresh("__bash_redir_coproc");
            let coproc_body = self.fresh("__bash_redir_coproc_body");
            let saved = self.fresh("__bash_saved");
            let coproc_output = self.fresh("__bash_coproc_output");
            let write_body = vec![
                let_stmt(&content, capture(out)),
                let_stmt(
                    &fd_key,
                    ternary(
                        method(ident(&target_var), "startsWith", vec![lit("/dev/fd/")]),
                        method(ident(&target_var), "slice", vec![int(8)]),
                        lit(""),
                    ),
                ),
                let_stmt(&coproc, index(ident("__bash_coprocs"), ident(&fd_key))),
                Statement::new(StmtKind::If {
                    cond: binary(BinOp::StrictNotEq, ident(&fd_key), lit("")),
                    then_body: vec![Statement::new(StmtKind::If {
                        cond: binary(BinOp::StrictEq, ident(&coproc), undefined()),
                        then_body: vec![
                            assign_stmt(
                                index(ident("__bash_fds"), ident(&fd_key)),
                                ident(&content),
                            ),
                            assign_stmt(ident("__bash_status"), int(0)),
                        ],
                        elifs: Vec::new(),
                        else_body: Some(vec![
                            let_stmt(&coproc_body, member(ident(&coproc), "body")),
                            let_stmt(&saved, ident("__bash_stdin")),
                            assign_stmt(ident("__bash_stdin"), ident(&content)),
                            let_stmt(
                                &coproc_output,
                                capture(vec![expr_stmt(call(
                                    ident(&coproc_body),
                                    vec![ident(&content)],
                                ))]),
                            ),
                            Statement::new(StmtKind::If {
                                cond: binary(
                                    BinOp::StrictEq,
                                    member(ident(&coproc), "stdout"),
                                    lit("parent"),
                                ),
                                then_body: vec![expr_stmt(bash_stdout_write_expr(ident(
                                    &coproc_output,
                                )))],
                                elifs: Vec::new(),
                                else_body: Some(vec![assign_stmt(
                                    index(ident("__bash_fds"), member(ident(&coproc), "read_fd")),
                                    ident(&coproc_output),
                                )]),
                            }),
                            assign_stmt(
                                index(ident("__bash_jobs"), member(ident(&coproc), "pid")),
                                ident("__bash_status"),
                            ),
                            assign_stmt(ident("__bash_stdin"), ident(&saved)),
                        ]),
                    })],
                    elifs: Vec::new(),
                    else_body: Some(vec![expr_stmt(call_named(
                        if append {
                            "__bash_append_file"
                        } else {
                            "__bash_write_file"
                        },
                        vec![bash_path_expr(ident(&target_var)), ident(&content)],
                    ))]),
                }),
            ];
            let wrapped = vec![
                let_stmt(&target_var, target),
                Statement::new(StmtKind::If {
                    cond: binary(BinOp::StrictEq, ident(&target_var), undefined()),
                    then_body: vec![assign_stmt(ident("__bash_status"), int(1))],
                    elifs: Vec::new(),
                    else_body: Some(write_body),
                }),
            ];
            out = wrapped;
        } else if let Some(target) = stdout_procsub {
            let content = self.fresh("__bash_procsub_input");
            let saved = self.fresh("__bash_saved");
            let mut body = procsub_body(target);
            rewrite_exit_to_return(&mut body);
            body.push(Statement::new(StmtKind::Return(Some(ident(
                "__bash_status",
            )))));
            let mut wrapped = vec![
                let_stmt(&content, capture(out)),
                let_stmt(&saved, ident("__bash_stdin")),
                assign_stmt(ident("__bash_stdin"), ident(&content)),
                expr_stmt(iife(body)),
                assign_stmt(ident("__bash_stdin"), ident(&saved)),
            ];
            out = std::mem::take(&mut wrapped);
        } else if let Some(target) = stdout_fd {
            let fd = self.fresh("__bash_write_fd");
            let content = self.fresh("__bash_write_fd_content");
            let coproc = self.fresh("__bash_write_fd_coproc");
            let coproc_body = self.fresh("__bash_write_fd_coproc_body");
            let saved = self.fresh("__bash_saved");
            let coproc_output = self.fresh("__bash_coproc_output");
            out = vec![
                let_stmt(&fd, param_value(target)),
                let_stmt(&content, capture(out)),
                let_stmt(&coproc, index(ident("__bash_coprocs"), ident(&fd))),
                Statement::new(StmtKind::If {
                    cond: binary(BinOp::StrictEq, ident(&coproc), undefined()),
                    then_body: vec![
                        assign_stmt(index(ident("__bash_fds"), ident(&fd)), ident(&content)),
                        assign_stmt(ident("__bash_status"), int(0)),
                    ],
                    elifs: Vec::new(),
                    else_body: Some(vec![
                        let_stmt(&coproc_body, member(ident(&coproc), "body")),
                        let_stmt(&saved, ident("__bash_stdin")),
                        assign_stmt(ident("__bash_stdin"), ident(&content)),
                        let_stmt(
                            &coproc_output,
                            capture(vec![expr_stmt(call(
                                ident(&coproc_body),
                                vec![ident(&content)],
                            ))]),
                        ),
                        Statement::new(StmtKind::If {
                            cond: binary(
                                BinOp::StrictEq,
                                member(ident(&coproc), "stdout"),
                                lit("parent"),
                            ),
                            then_body: vec![expr_stmt(bash_stdout_write_expr(ident(
                                &coproc_output,
                            )))],
                            elifs: Vec::new(),
                            else_body: Some(vec![assign_stmt(
                                index(ident("__bash_fds"), member(ident(&coproc), "read_fd")),
                                ident(&coproc_output),
                            )]),
                        }),
                        assign_stmt(
                            index(ident("__bash_jobs"), member(ident(&coproc), "pid")),
                            ident("__bash_status"),
                        ),
                        assign_stmt(ident("__bash_stdin"), ident(&saved)),
                    ]),
                }),
            ];
        }
        if stdout_to_stderr {
            let saved = self.fresh("__bash_saved_stdout_to_stderr");
            let mut wrapped = vec![
                let_stmt(&saved, ident("__bash_stdout_to_stderr")),
                assign_stmt(ident("__bash_stdout_to_stderr"), Expression::bool(true)),
            ];
            wrapped.extend(out);
            wrapped.push(assign_stmt(ident("__bash_stdout_to_stderr"), ident(&saved)));
            out = wrapped;
        }
        if stderr_to_stdout.is_some() || stderr_null.is_some() {
            let saved_to_stdout = self.fresh("__bash_saved_stderr_to_stdout");
            let saved_null = self.fresh("__bash_saved_stderr_null");
            let mut wrapped = vec![
                let_stmt(&saved_to_stdout, ident("__bash_stderr_to_stdout")),
                let_stmt(&saved_null, ident("__bash_stderr_null")),
            ];
            if let Some(value) = stderr_to_stdout {
                wrapped.push(assign_stmt(
                    ident("__bash_stderr_to_stdout"),
                    Expression::bool(value),
                ));
            }
            if let Some(value) = stderr_null {
                wrapped.push(assign_stmt(
                    ident("__bash_stderr_null"),
                    Expression::bool(value),
                ));
            }
            wrapped.extend(out);
            wrapped.push(assign_stmt(ident("__bash_stderr_null"), ident(&saved_null)));
            wrapped.push(assign_stmt(
                ident("__bash_stderr_to_stdout"),
                ident(&saved_to_stdout),
            ));
            out = wrapped;
        }
        if let Some(content) = stdin {
            let saved = self.fresh("__bash_saved");
            let mut wrapped = vec![
                let_stmt(&saved, ident("__bash_stdin")),
                assign_stmt(ident("__bash_stdin"), content),
            ];
            wrapped.extend(out);
            wrapped.push(assign_stmt(ident("__bash_stdin"), ident(&saved)));
            out = wrapped;
        }
        Ok(out)
    }

    fn persistent_redirections(&mut self, redirs: &[Pair<Rule>]) -> R<Vec<Statement>> {
        self.persistent_redirections_with_fd_override(redirs, None)
    }

    fn persistent_redirections_with_fd_override(
        &mut self,
        redirs: &[Pair<Rule>],
        fd_override: Option<String>,
    ) -> R<Vec<Statement>> {
        let mut out = Vec::new();
        for r in redirs {
            let (mut fd, op, target, mut before) = self.redirection_parts(r.clone())?;
            out.append(&mut before);
            if fd == "0" && matches!(op.as_str(), "<<<" | "<<" | "<") {
                if let Some(override_fd) = &fd_override {
                    fd = override_fd.clone();
                }
            }
            match op.as_str() {
                "<<<" => {
                    out.push(self.assign_fd_input(
                        &fd,
                        parts_to_expr(vec![Part::Expr(target), Part::Text("\n".into())]),
                    ));
                    out.push(assign_stmt(ident("__bash_status"), int(0)));
                }
                "<<" => {
                    out.push(self.assign_fd_input(&fd, target));
                    out.push(assign_stmt(ident("__bash_status"), int(0)));
                }
                "<" => {
                    let content = match target.kind {
                        ExprKind::Call { .. } if is_procsub(&target) => procsub_capture(target),
                        _ => self.read_file_or_fd_expr(target),
                    };
                    out.push(self.assign_fd_input(&fd, content));
                }
                "<&" if literal_string(&target) == Some("-") => {
                    out.push(self.assign_fd_input(&fd, undefined()));
                    out.push(assign_stmt(ident("__bash_status"), int(0)));
                }
                "<&" => {
                    let content = self.read_fd_expr(target);
                    out.push(self.assign_fd_input(&fd, content));
                }
                _ => {
                    out.push(assign_stmt(ident("__bash_status"), int(0)));
                }
            }
        }
        if out.is_empty() {
            out.push(assign_stmt(ident("__bash_status"), int(0)));
        }
        Ok(out)
    }

    fn assign_fd_input(&self, fd: &str, content: Expression) -> Statement {
        if fd == "0" {
            assign_stmt(ident("__bash_stdin"), content)
        } else {
            assign_stmt(index(ident("__bash_fds"), lit(fd)), content)
        }
    }

    fn read_file_or_fd_expr(&mut self, target: Expression) -> Expression {
        let raw = self.fresh("__bash_read_raw_path");
        let path = self.fresh("__bash_read_path");
        let key = self.fresh("__bash_fd_key");
        let value = self.fresh("__bash_fd_value");
        sequence(vec![
            assign_expr(ident(&raw), target),
            assign_expr(
                ident(&path),
                ternary(
                    method(ident(&raw), "startsWith", vec![lit("/")]),
                    ident(&raw),
                    parts_to_expr(vec![
                        Part::Expr(ident("PWD")),
                        Part::Text("/".into()),
                        Part::Expr(ident(&raw)),
                    ]),
                ),
            ),
            assign_expr(ident(&key), lit("")),
            assign_expr(ident(&value), undefined()),
            ternary(
                method(ident(&path), "startsWith", vec![lit("/dev/fd/")]),
                sequence(vec![
                    assign_expr(ident(&key), method(ident(&path), "slice", vec![int(8)])),
                    assign_expr(ident(&value), index(ident("__bash_fds"), ident(&key))),
                    ternary(
                        binary(BinOp::StrictEq, ident(&value), undefined()),
                        sequence(vec![assign_expr(ident("__bash_status"), int(1)), lit("")]),
                        sequence(vec![
                            assign_expr(ident("__bash_status"), int(0)),
                            ident(&value),
                        ]),
                    ),
                ]),
                ternary(
                    call_named("__bash_file_exists", vec![ident(&path)]),
                    sequence(vec![
                        assign_expr(ident("__bash_status"), int(0)),
                        call_named("__bash_read_file", vec![ident(&path)]),
                    ]),
                    sequence(vec![assign_expr(ident("__bash_status"), int(1)), lit("")]),
                ),
            ),
        ])
    }

    fn read_fd_expr(&mut self, target: Expression) -> Expression {
        let key = self.fresh("__bash_dup_fd");
        let value = self.fresh("__bash_dup_fd_value");
        iife(vec![
            let_stmt(&key, param_value(target)),
            let_stmt(&value, index(ident("__bash_fds"), ident(&key))),
            Statement::new(StmtKind::Return(Some(ternary(
                binary(BinOp::StrictEq, ident(&value), undefined()),
                sequence(vec![assign_expr(ident("__bash_status"), int(1)), lit("")]),
                sequence(vec![
                    assign_expr(ident("__bash_status"), int(0)),
                    ident(&value),
                ]),
            )))),
        ])
    }

    fn bash_string_slice_expr(
        &mut self,
        target: Expression,
        offset: Expression,
        length: Option<Expression>,
    ) -> Expression {
        let value = self.fresh("__bash_slice_value");
        let start = self.fresh("__bash_slice_start");
        let end = self.fresh("__bash_slice_end");
        let check_negative_end = length.as_ref().is_some_and(expr_is_negative_literal);
        let target_value = param_value(target);
        let start_expr = ternary(
            binary(BinOp::Lt, offset.clone(), int(0)),
            binary(BinOp::Add, member(ident(&value), "length"), offset.clone()),
            offset,
        );
        let end_expr = match length {
            None => member(ident(&value), "length"),
            Some(len) => ternary(
                binary(BinOp::Lt, len.clone(), int(0)),
                binary(BinOp::Add, member(ident(&value), "length"), len.clone()),
                binary(BinOp::Add, ident(&start), len),
            ),
        };
        iife(vec![
            let_stmt(&value, target_value),
            let_stmt(&start, start_expr),
            let_stmt(&end, end_expr),
            Statement::new(StmtKind::If {
                cond: binary(BinOp::Lt, ident(&start), int(0)),
                then_body: vec![Statement::new(StmtKind::Return(Some(lit(""))))],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::If {
                cond: binary(
                    BinOp::And,
                    Expression::bool(check_negative_end),
                    binary(BinOp::Lt, ident(&end), ident(&start)),
                ),
                then_body: vec![
                    bash_stderr_stmt(lit("bash: substring expression < 0\n")),
                    assign_stmt(ident("__bash_status"), int(1)),
                    Statement::new(StmtKind::Return(Some(lit("")))),
                ],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::Return(Some(method(
                ident(&value),
                "slice",
                vec![ident(&start), ident(&end)],
            )))),
        ])
    }

    fn redirection_parts(
        &mut self,
        pair: Pair<Rule>,
    ) -> R<(String, String, Expression, Vec<Statement>)> {
        let raw = pair.as_str().to_string();
        let inner = pair.into_inner().next().ok_or("empty redirection")?;
        match inner.as_rule() {
            Rule::here_string => {
                let mut fd = "0".to_string();
                let mut word = None;
                if let Some(raw_fd) = here_string_raw_fd(&raw) {
                    fd = raw_fd;
                }
                for p in inner.into_inner() {
                    match p.as_rule() {
                        Rule::redir_fd => fd = p.as_str().to_string(),
                        Rule::word => word = Some(p),
                        _ => {}
                    }
                }
                let w = word.ok_or("here-string without word")?;
                Ok((fd, "<<<".into(), self.here_string_word_expr(w)?, Vec::new()))
            }
            Rule::here_document => {
                let mut fd = "0".to_string();
                for p in inner.clone().into_inner() {
                    if p.as_rule() == Rule::redir_fd {
                        fd = p.as_str().to_string();
                        break;
                    }
                }
                let doc = self
                    .heredocs
                    .pop_front()
                    .ok_or("here-document body missing")?;
                let body = if doc.expand {
                    self.heredoc_body_expr(&doc.body)?
                } else {
                    lit(&doc.body)
                };
                let before = if let Some(warning) = doc.eof_warning {
                    vec![bash_stderr_stmt(lit(&warning))]
                } else {
                    Vec::new()
                };
                Ok((fd, "<<".into(), body, before))
            }
            Rule::fd_redirection => {
                let mut fd = String::new();
                let mut op = String::new();
                let mut target = lit("");
                for p in inner.into_inner() {
                    match p.as_rule() {
                        Rule::redir_fd => fd = p.as_str().to_string(),
                        Rule::redir_op => op = p.as_str().to_string(),
                        Rule::redir_target => {
                            let t = p.into_inner().next().ok_or("empty redirection target")?;
                            target = match t.as_rule() {
                                Rule::redir_close => lit("-"),
                                _ => self.redirection_target_expr(t, &op)?,
                            };
                        }
                        _ => {}
                    }
                }
                if fd.is_empty() {
                    fd = if op.starts_with('<') {
                        "0".into()
                    } else {
                        "1".into()
                    };
                }
                Ok((fd, op, target, Vec::new()))
            }
            other => Err(format!("unsupported redirection {other:?}")),
        }
    }

    fn redirection_target_expr(&mut self, target: Pair<Rule>, op: &str) -> R<Expression> {
        let mut fallback = self.word_expr(target.clone())?;
        if !word_has_quoted_part(&target) && word_has_unquoted_expansion(&target) {
            let raw = self.fresh("__bash_redir_raw");
            let fields = self.fresh("__bash_redir_fields");
            fallback = sequence(vec![
                assign_expr(ident(&raw), fallback),
                assign_expr(ident(&fields), split_bash_words(ident(&raw))),
                ternary(
                    binary(BinOp::StrictNotEq, member(ident(&fields), "length"), int(1)),
                    sequence(vec![
                        bash_stderr_write_expr(call_named(
                            "__bash_sprintf",
                            vec![
                                lit("bash: %s: ambiguous redirect\n"),
                                call_named("__bash_string", vec![ident(&raw)]),
                            ],
                        )),
                        undefined(),
                    ]),
                    ident(&raw),
                ),
            ]);
        }
        if self.shell_flags.contains(&'f')
            || !(op == ">" || op == ">|" || op == ">>" || op == "&>" || op == "&>>")
        {
            return Ok(fallback);
        }
        let Some(text) = glob_word_text(&target) else {
            return Ok(fallback);
        };
        if !word_text_has_glob_meta(&text) {
            return Ok(fallback);
        }
        let Some(matches) = self.path_glob_matches_expr(&text) else {
            return Ok(fallback);
        };
        let tmp = self.fresh("__bash_redir_matches");
        Ok(iife_with_args(
            vec![
                let_stmt(&tmp, matches),
                Statement::new(StmtKind::Return(Some(ternary(
                    binary(BinOp::Gt, member(ident(&tmp), "length"), int(1)),
                    sequence(vec![
                        bash_stderr_write_expr(call_named(
                            "__bash_sprintf",
                            vec![lit("bash: %s: ambiguous redirect\n"), lit(&text)],
                        )),
                        undefined(),
                    ]),
                    ternary(
                        binary(BinOp::StrictEq, member(ident(&tmp), "length"), int(1)),
                        index(ident(&tmp), int(0)),
                        lit(&text),
                    ),
                )))),
            ],
            vec![
                param_named("PWD"),
                param_named("__bash_stderr_null"),
                param_named("__bash_stderr_to_stdout"),
            ],
            vec![
                ident("PWD"),
                ident("__bash_stderr_null"),
                ident("__bash_stderr_to_stdout"),
            ],
        ))
    }

    fn heredoc_body_expr(&mut self, body: &str) -> R<Expression> {
        let pairs = WashmParser::parse(Rule::heredoc_body, body)
            .map_err(|e| format!("here-document: {e}"))?;
        let mut parts = Vec::new();
        for p in pairs {
            if p.as_rule() == Rule::heredoc_body {
                for part in p.into_inner() {
                    match part.as_rule() {
                        Rule::hd_text => parts.push(Part::Text(part.as_str().to_string())),
                        Rule::hd_escape => parts.push(Part::Text(part.as_str()[1..].to_string())),
                        _ => parts.push(Part::Expr(self.part_expr(part)?)),
                    }
                }
            }
        }
        Ok(parts_to_expr(parts))
    }

    // ── assignments ──────────────────────────────────────────────────────

    fn walk_assignment(&mut self, pair: Pair<Rule>) -> R<Statement> {
        let (target, op, mut value) = self.assignment_parts(pair)?;
        match &target {
            AssignTarget::Name(name) | AssignTarget::Element(name, _) => {
                if is_readonly_parameter_target(name) || self.target_text_is_readonly(name) {
                    return Ok(assign_stmt(ident("__bash_status"), int(1)));
                }
            }
        }
        let target_expr = match target {
            AssignTarget::Name(n) => {
                let resolved = self.resolve_nameref_text(&n).unwrap_or_else(|| n.clone());
                self.record_variable(&resolved);
                if resolved != n {
                    self.record_variable(&n);
                }
                if op != "+=" {
                    value = self.apply_variable_attributes(&resolved, value)?;
                }
                if op != "+=" && is_name(&resolved) {
                    if let Some(snapshot) =
                        bash_array_snapshot(&value, self.assoc_arrays.contains(&resolved))
                    {
                        self.array_values.insert(resolved.clone(), snapshot);
                        if self.assoc_arrays.contains(&resolved) {
                            self.indexed_arrays.remove(&resolved);
                        } else {
                            self.indexed_arrays.insert(resolved.clone());
                        }
                        self.variable_values.remove(&resolved);
                    } else if let Some(s) = literal_string(&value) {
                        self.variable_values.insert(resolved.clone(), s.to_string());
                        self.array_values.remove(&resolved);
                    } else {
                        self.variable_values.remove(&resolved);
                        self.array_values.remove(&resolved);
                    }
                } else {
                    self.variable_values.remove(&n);
                    self.array_values.remove(&n);
                }
                if self.shell_flags.contains(&'a') {
                    self.exported_vars.insert(resolved.clone());
                }
                let target = self.name_or_element_target(&n)?;
                target
            }
            AssignTarget::Element(arr, key) => {
                let target = self
                    .resolve_nameref_text(&arr)
                    .unwrap_or_else(|| arr.clone());
                if let (Some(key_text), Some(value_text)) = (
                    literal_key_string(&key),
                    literal_string(&value).map(str::to_string),
                ) {
                    self.record_variable(&target);
                    if !self.assoc_arrays.contains(&target) {
                        self.indexed_arrays.insert(target.clone());
                    }
                    self.record_array_element_value(&target, key_text, value_text);
                } else {
                    self.array_values.remove(&target);
                }
                if target.contains('[') {
                    self.name_or_element_target(&target)?
                } else {
                    index(ident(&target), key)
                }
            }
        };
        if op == "+=" {
            if let ExprKind::Ident(name) = &target_expr.kind {
                if self.integer_vars.contains(name) {
                    return Ok(assign_stmt(
                        target_expr.clone(),
                        call_named(
                            "__bash_string",
                            vec![binary(BinOp::Add, to_number(target_expr), to_number(value))],
                        ),
                    ));
                }
            }
            return Ok(compound_append(target_expr, value));
        }
        Ok(assign_stmt(target_expr, value))
    }

    /// `name[sub]` `op` `value` of an `assignment_word`.
    fn assignment_parts(&mut self, pair: Pair<Rule>) -> R<(AssignTarget, String, Expression)> {
        self.assignment_parts_with_array_kind(pair, false)
    }

    fn assignment_word_static_rhs_text(&mut self, pair: &Pair<Rule>) -> Option<String> {
        pair.clone().into_inner().find_map(|p| match p.as_rule() {
            Rule::word | Rule::assignment_value_word => self.static_word_text(&p),
            _ => None,
        })
    }

    fn assignment_parts_with_array_kind(
        &mut self,
        pair: Pair<Rule>,
        force_assoc_array: bool,
    ) -> R<(AssignTarget, String, Expression)> {
        let mut target = None;
        let mut target_name = None;
        let mut op = String::from("=");
        let mut value: Option<Expression> = None;
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::assign_target => {
                    let mut name = String::new();
                    let mut sub = None;
                    let mut sub_raw = None;
                    for t in p.into_inner() {
                        match t.as_rule() {
                            Rule::name => name = t.as_str().to_string(),
                            Rule::subscript => {
                                sub_raw = Some(t.as_str().to_string());
                                sub = Some(self.subscript_expr(t)?);
                            }
                            _ => {}
                        }
                    }
                    target_name = Some(name.clone());
                    target = Some(match sub {
                        Some(Subscript::Key(k)) => {
                            let key = if self.assoc_arrays.contains(&name) {
                                sub_raw
                                    .as_deref()
                                    .map(|raw| {
                                        raw.strip_prefix('[')
                                            .and_then(|s| s.strip_suffix(']'))
                                            .unwrap_or(raw)
                                    })
                                    .map(|raw| self.subscript_assoc_key_expr(raw))
                                    .transpose()?
                                    .unwrap_or(k)
                            } else {
                                k
                            };
                            AssignTarget::Element(name, key)
                        }
                        Some(Subscript::All) | None => AssignTarget::Name(name),
                    });
                }
                Rule::assign_op => op = p.as_str().to_string(),
                Rule::array_literal => {
                    let as_assoc = force_assoc_array
                        || target_name
                            .as_ref()
                            .is_some_and(|name| self.assoc_arrays.contains(name));
                    value = Some(self.array_literal_expr(p, as_assoc)?);
                }
                Rule::word | Rule::assignment_value_word => {
                    value = Some(if pair_has_command_substitution(p.clone())
                        || (self.function_depth > 0 && !self.inline_call_context)
                        || self.dynamic_arith
                    {
                        self.assignment_word_value_expr(p)?
                    } else {
                        match self.static_word_text(&p) {
                            Some(text) => lit(&text),
                            None => self.assignment_word_value_expr(p)?,
                        }
                    });
                }
                _ => {}
            }
        }
        Ok((
            target.ok_or("assignment without a target")?,
            op,
            value.unwrap_or_else(|| lit("")),
        ))
    }

    fn assignment_word_value_expr(&mut self, pair: Pair<Rule>) -> R<Expression> {
        if let Some(e) = assignment_tilde_literal_expr(pair.as_str()) {
            return Ok(e);
        }
        let parts = self.word_parts(pair)?;
        Ok(parts_to_expr(expand_assignment_tilde_after_colon(parts)))
    }

    /// `( a b [k]=v … )` — an indexed array, or a map when any element has a key.
    fn array_literal_expr(&mut self, pair: Pair<Rule>, as_assoc: bool) -> R<Expression> {
        if !as_assoc
            && !pair
                .clone()
                .into_inner()
                .any(|el| el.as_rule() == Rule::associative_element)
        {
            let mut args = Vec::new();
            for el in pair.into_inner() {
                if el.as_rule() == Rule::word {
                    args.extend(self.words_shell_args(vec![el])?);
                }
            }
            return Ok(shell_args_array(args));
        }

        let mut plain = Vec::new();
        let mut keyed: Vec<(Expression, Expression)> = Vec::new();
        let mut indexed_keyed: Vec<(Option<i64>, Expression)> = Vec::new();
        for el in pair.into_inner() {
            match el.as_rule() {
                Rule::word => {
                    for value in self.expand_word(el)? {
                        if as_assoc {
                            plain.push(value);
                        } else {
                            indexed_keyed.push((None, value));
                        }
                    }
                }
                Rule::associative_element => {
                    let mut key = None;
                    let mut key_raw = None;
                    let mut val = lit("");
                    for p in el.into_inner() {
                        match p.as_rule() {
                            Rule::subscript => {
                                key_raw = Some(p.as_str().to_string());
                                key = Some(self.subscript_expr(p)?);
                            }
                            Rule::word => val = self.word_expr(p)?,
                            _ => {}
                        }
                    }
                    if as_assoc {
                        let key = match key {
                            Some(Subscript::Key(k)) => key_raw
                                .as_deref()
                                .map(|raw| {
                                    raw.strip_prefix('[')
                                        .and_then(|s| s.strip_suffix(']'))
                                        .unwrap_or(raw)
                                })
                                .map(|raw| self.subscript_assoc_key_expr(raw))
                                .transpose()?
                                .unwrap_or(k),
                            _ => lit(""),
                        };
                        keyed.push((key, val));
                    } else {
                        let idx = key_raw.as_deref().and_then(|raw| {
                            raw.strip_prefix('[')
                                .and_then(|s| s.strip_suffix(']'))
                                .unwrap_or(raw)
                                .parse::<i64>()
                                .ok()
                        });
                        indexed_keyed.push((idx, val));
                    }
                }
                _ => {}
            }
        }
        if !as_assoc {
            let mut next = 0_i64;
            let mut slots: Vec<Expression> = Vec::new();
            for (idx, value) in indexed_keyed {
                let idx = idx.unwrap_or(next).max(0);
                while slots.len() <= idx as usize {
                    slots.push(undefined());
                }
                slots[idx as usize] = value;
                next = idx + 1;
            }
            return Ok(array(slots));
        }
        if keyed.is_empty() {
            return Ok(array(plain));
        }
        Ok(Expression::new(ExprKind::Object(
            keyed
                .into_iter()
                .map(|(key, value)| ObjectProperty::Computed { key, value })
                .collect(),
        )))
    }

    // ── functions and control flow ───────────────────────────────────────

    fn walk_function_def(&mut self, pair: Pair<Rule>) -> R<Statement> {
        let display_source = pair.as_str().trim().to_string();
        let parts: Vec<Pair<Rule>> = pair.into_inner().collect();
        let name = parts
            .iter()
            .find(|p| p.as_rule() == Rule::function_name)
            .map(|p| p.as_str().to_string())
            .unwrap_or_default();
        let emitted_name = if name.is_empty() {
            name.clone()
        } else {
            let emitted = self.fresh("__bash_fn");
            self.functions.insert(name.clone(), emitted.clone());
            self.function_bodies.insert(name.clone(), display_source);
            emitted
        };
        let mut body = Vec::new();
        for inner in parts {
            match inner.as_rule() {
                Rule::function_name => {}
                Rule::function_body => {
                    let command_source = inner.as_str().trim().to_string();
                    let mut redirs = Vec::new();
                    let mut cmd = None;
                    for b in inner.into_inner() {
                        match b.as_rule() {
                            Rule::redirection => redirs.push(b),
                            _ => cmd = Some(b),
                        }
                    }
                    let saved_variable_values = self.variable_values.clone();
                    let saved_array_values = self.array_values.clone();
                    let saved_indexed_arrays = self.indexed_arrays.clone();
                    let saved_assoc_arrays = self.assoc_arrays.clone();
                    let saved_readonly_vars = self.readonly_vars.clone();
                    let saved_integer_vars = self.integer_vars.clone();
                    let saved_lowercase_vars = self.lowercase_vars.clone();
                    let saved_uppercase_vars = self.uppercase_vars.clone();
                    let saved_namerefs = self.namerefs.clone();
                    let saved_shell_flags = self.shell_flags.clone();
                    let saved_dynamic_arith = self.dynamic_arith;
                    self.function_depth += 1;
                    self.dynamic_arith = true;
                    let stmts_result: R<Vec<Statement>> = match cmd {
                        Some(c) => self.command_stmts(c),
                        None => Ok(Vec::new()),
                    };
                    self.function_depth = self.function_depth.saturating_sub(1);
                    self.dynamic_arith = saved_dynamic_arith;
                    self.variable_values = saved_variable_values;
                    self.array_values = saved_array_values;
                    self.indexed_arrays = saved_indexed_arrays;
                    self.assoc_arrays = saved_assoc_arrays;
                    self.readonly_vars = saved_readonly_vars;
                    self.integer_vars = saved_integer_vars;
                    self.lowercase_vars = saved_lowercase_vars;
                    self.uppercase_vars = saved_uppercase_vars;
                    self.namerefs = saved_namerefs;
                    self.shell_flags = saved_shell_flags;
                    let stmts = stmts_result?;
                    body = if redirs.is_empty() {
                        stmts
                    } else {
                        self.apply_redirections(stmts, &redirs)?
                    };
                    if function_source_has_local_dash(&command_source) {
                        let tmp = self.fresh("__bash_local_flags");
                        body.insert(0, let_stmt(&tmp, ident("__bash_flags")));
                        body.push(assign_stmt(ident("__bash_flags"), ident(&tmp)));
                    }
                    if !name.is_empty() && redirs.is_empty() {
                        self.function_commands.insert(name.clone(), command_source);
                    } else if !name.is_empty() {
                        self.function_commands.remove(&name);
                    }
                }
                _ => {}
            }
        }
        let mut locals = Vec::new();
        collect_function_scoped_names(&body, &mut locals);
        rewrite_function_scoped_decls(&mut body);
        for name in locals.into_iter().rev() {
            body.insert(0, let_stmt(&name, undefined()));
        }
        // A function without an explicit return yields the status of its
        // last command.
        if !matches!(body.last().map(|s| &s.kind), Some(StmtKind::Return(_))) {
            body.extend(funcname_pop_exprs().into_iter().map(expr_stmt));
            body.push(Statement::new(StmtKind::Return(Some(ident("__bash_status")))));
        }
        for stmt in funcname_push_stmts(&name).into_iter().rev() {
            body.insert(0, stmt);
        }
        let stmt = Statement::new(StmtKind::FunctionDecl {
            name: emitted_name,
            params: bash_function_params(),
            return_type: None,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        });
        if self.function_depth > 0 {
            self.hoisted_functions.push(stmt);
            Ok(Statement::new(StmtKind::Empty))
        } else {
            Ok(stmt)
        }
    }

    fn walk_if(&mut self, pair: Pair<Rule>) -> R<Statement> {
        if pair.as_str().contains("elif") {
            if let Some(stmt) = self.walk_if_from_source(pair.as_str())? {
                return Ok(stmt);
            }
        }
        let mut cond = None;
        let mut then_body = Vec::new();
        let mut elifs: Vec<(Expression, Vec<Statement>)> = Vec::new();
        let mut else_body = None;
        let mut seen_then = false;
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::kw_then => seen_then = true,
                Rule::list => {
                    if !seen_then {
                        cond = Some(self.list_expr(inner)?.cond());
                    } else {
                        then_body = self.walk_list(inner)?;
                    }
                }
                Rule::elif_clause => {
                    let mut c = None;
                    let mut b = Vec::new();
                    let mut after_then = false;
                    for p in inner.into_inner() {
                        match p.as_rule() {
                            Rule::kw_then => after_then = true,
                            Rule::list if !after_then => c = Some(self.list_expr(p)?.cond()),
                            Rule::list => b = self.walk_list(p)?,
                            _ => {}
                        }
                    }
                    elifs.push((c.ok_or("elif without condition")?, b));
                }
                Rule::else_clause => {
                    let mut b = Vec::new();
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::list {
                            b = self.walk_list(p)?;
                        }
                    }
                    else_body = Some(b);
                }
                _ => {}
            }
        }
        Ok(Statement::new(StmtKind::If {
            cond: cond.ok_or("if without condition")?,
            then_body,
            elifs,
            else_body,
        }))
    }

    fn walk_if_from_source(&mut self, source: &str) -> R<Option<Statement>> {
        let Some(mut rest) = source.trim().strip_prefix("if") else {
            return Ok(None);
        };
        rest = rest.trim_start();
        let Some(then_pos) = find_control_keyword(rest, "then") else {
            return Ok(None);
        };
        let cond_src = trim_shell_fragment(&rest[..then_pos]);
        rest = rest[then_pos + "then".len()..].trim_start();

        let mut cond = Some(self.list_expr_from_source(cond_src)?.cond());
        let mut then_body = Vec::new();
        let mut elifs: Vec<(Expression, Vec<Statement>)> = Vec::new();
        let mut else_body = None;

        loop {
            let Some((kw, pos)) = find_next_if_boundary(rest) else {
                return Ok(None);
            };
            let body_src = trim_shell_fragment(&rest[..pos]);
            match kw {
                "elif" => {
                    if then_body.is_empty() {
                        then_body = self.walk_list_from_source(body_src)?;
                    } else {
                        let last = elifs.last_mut().ok_or("elif body without condition")?;
                        last.1 = self.walk_list_from_source(body_src)?;
                    }
                    rest = rest[pos + kw.len()..].trim_start();
                    let Some(next_then) = find_control_keyword(rest, "then") else {
                        return Ok(None);
                    };
                    let elif_cond_src = trim_shell_fragment(&rest[..next_then]);
                    let elif_cond = self.list_expr_from_source(elif_cond_src)?.cond();
                    elifs.push((elif_cond, Vec::new()));
                    rest = rest[next_then + "then".len()..].trim_start();
                }
                "else" => {
                    if then_body.is_empty() {
                        then_body = self.walk_list_from_source(body_src)?;
                    } else {
                        let last = elifs.last_mut().ok_or("else after empty elif")?;
                        last.1 = self.walk_list_from_source(body_src)?;
                    }
                    rest = rest[pos + kw.len()..].trim_start();
                    let Some(fi_pos) = find_control_keyword(rest, "fi") else {
                        return Ok(None);
                    };
                    else_body = Some(self.walk_list_from_source(trim_shell_fragment(
                        &rest[..fi_pos],
                    ))?);
                    break;
                }
                "fi" => {
                    if then_body.is_empty() {
                        then_body = self.walk_list_from_source(body_src)?;
                    } else if let Some(last) = elifs.last_mut() {
                        last.1 = self.walk_list_from_source(body_src)?;
                    }
                    break;
                }
                _ => return Ok(None),
            }
        }

        Ok(Some(Statement::new(StmtKind::If {
            cond: cond.take().ok_or("if without condition")?,
            then_body,
            elifs,
            else_body,
        })))
    }

    fn walk_list_from_source(&mut self, source: &str) -> R<Vec<Statement>> {
        let mut pairs = WashmParser::parse(Rule::list, source)
            .map_err(|e| format!("if list fragment: {e}"))?;
        let pair = pairs.next().ok_or("empty if list fragment")?;
        self.walk_list(pair)
    }

    fn list_expr_from_source(&mut self, source: &str) -> R<Cmd> {
        let mut pairs = WashmParser::parse(Rule::list, source)
            .map_err(|e| format!("if condition fragment: {e}"))?;
        let pair = pairs.next().ok_or("empty if condition fragment")?;
        self.list_expr(pair)
    }

    fn walk_while(&mut self, pair: Pair<Rule>, until: bool) -> R<Statement> {
        let mut cond = None;
        let mut body = Vec::new();
        let mut seen_do = false;
        let saved_dynamic_arith = self.dynamic_arith;
        self.dynamic_arith = true;
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::kw_do => seen_do = true,
                Rule::list if !seen_do => {
                    let saved_variable_values = self.variable_values.clone();
                    let saved_array_values = self.array_values.clone();
                    self.variable_values.clear();
                    self.array_values.clear();
                    cond = Some(self.list_expr(inner)?.cond());
                    self.variable_values = saved_variable_values;
                    self.array_values = saved_array_values;
                }
                Rule::list => body = self.walk_list(inner)?,
                _ => {}
            }
        }
        self.dynamic_arith = saved_dynamic_arith;
        let mut cond = cond.ok_or("loop without condition")?;
        if until {
            cond = unary(UnaryOp::Not, cond);
        }
        let ran = self.fresh("__bash_while_ran");
        let body_status = self.fresh("__bash_while_body_status");
        let mut wrapped_body = vec![assign_stmt(ident(&ran), Expression::bool(true))];
        wrapped_body.extend(body);
        wrapped_body.push(assign_stmt(ident(&body_status), ident("__bash_status")));
        Ok(Statement::new(StmtKind::Block(vec![
            let_stmt(&ran, Expression::bool(false)),
            let_stmt(&body_status, int(0)),
            assign_stmt(ident("__bash_status"), int(0)),
            Statement::new(StmtKind::While {
                cond,
                body: wrapped_body,
                else_body: None,
            }),
            assign_stmt(
                ident("__bash_status"),
                ternary(ident(&ran), ident(&body_status), int(0)),
            ),
        ])))
    }

    fn walk_while_read_with_redirection(
        &mut self,
        pair: Pair<Rule>,
        redirs: &[Pair<Rule>],
    ) -> R<Option<Lowered>> {
        let mut cond_list = None;
        let mut body_list = None;
        let mut seen_do = false;
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::kw_do => seen_do = true,
                Rule::list if !seen_do => cond_list = Some(inner),
                Rule::list => body_list = Some(inner),
                _ => {}
            }
        }
        let Some(cond_list) = cond_list else {
            return Ok(None);
        };
        let Some((name, suffix)) = self.single_simple_command_parts(cond_list)? else {
            return Ok(None);
        };
        if name != "read" {
            return Ok(None);
        }
        let Some((input, mut before_stmts)) = self.stdin_expr_from_redirs(redirs)? else {
            return Ok(None);
        };
        let saved_dynamic_arith = self.dynamic_arith;
        self.dynamic_arith = true;
        let body = match body_list {
            Some(body) => self.walk_list(body)?,
            None => Vec::new(),
        };
        self.dynamic_arith = saved_dynamic_arith;

        let item = self.fresh("__bash_read_line");
        let saved_stdin = self.fresh("__bash_saved_stdin");
        let mut loop_body = vec![assign_stmt(
            ident("__bash_stdin"),
            parts_to_expr(vec![Part::Expr(ident(&item)), Part::Text("\n".into())]),
        )];
        loop_body.extend(self.walk_read(suffix.clone(), false)?.into_stmts());
        loop_body.extend(body);
        let mut eof_read = vec![assign_stmt(ident("__bash_stdin"), lit(""))];
        eof_read.extend(self.walk_read(suffix, false)?.into_stmts());

        before_stmts.extend(vec![
            let_stmt(&saved_stdin, ident("__bash_stdin")),
            Statement::new(StmtKind::ForIn {
                var: item,
                key: None,
                iter: bash_input_lines(input),
                body: loop_body,
                of: true,
                else_body: None,
                is_async: false,
            }),
            Statement::new(StmtKind::Block(eof_read)),
            assign_stmt(ident("__bash_stdin"), ident(&saved_stdin)),
            assign_stmt(ident("__bash_status"), int(0)),
        ]);
        Ok(Some(Lowered::stmts(before_stmts)))
    }

    fn single_simple_command_parts<'a>(
        &mut self,
        list: Pair<'a, Rule>,
    ) -> R<Option<(String, Vec<Pair<'a, Rule>>)>> {
        let and_ors = list
            .into_inner()
            .filter(|p| p.as_rule() == Rule::and_or)
            .collect::<Vec<_>>();
        if and_ors.len() != 1 {
            return Ok(None);
        }
        let pipelines = and_ors[0]
            .clone()
            .into_inner()
            .filter(|p| p.as_rule() == Rule::pipeline)
            .collect::<Vec<_>>();
        if pipelines.len() != 1 {
            return Ok(None);
        }
        let (_, _, units) = self.pipeline_units(pipelines[0].clone());
        if units.len() != 1 || !units[0].redirs.is_empty() {
            return Ok(None);
        }
        Ok(simple_command_literal_parts(units[0].cmd.clone()))
    }

    fn stdin_expr_from_redirs(
        &mut self,
        redirs: &[Pair<Rule>],
    ) -> R<Option<(Expression, Vec<Statement>)>> {
        let mut stdin = None;
        let mut before_stmts = Vec::new();
        for r in redirs {
            let (fd, op, target, mut before) = self.redirection_parts(r.clone())?;
            before_stmts.append(&mut before);
            match op.as_str() {
                "<<<" if fd == "0" => {
                    stdin = Some(parts_to_expr(vec![
                        Part::Expr(target),
                        Part::Text("\n".into()),
                    ]));
                }
                "<<" if fd == "0" => stdin = Some(target),
                "<" if fd == "0" => {
                    stdin = Some(match target.kind {
                        ExprKind::Call { .. } if is_procsub(&target) => procsub_capture(target),
                        _ => self.read_file_or_fd_expr(target),
                    });
                }
                "<&" if fd == "0" && literal_string(&target) != Some("-") => {
                    stdin = Some(self.read_fd_expr(target));
                }
                _ => {}
            }
        }
        Ok(stdin.map(|input| (input, before_stmts)))
    }

    fn loop_body(&mut self, pair: Pair<Rule>) -> R<Vec<Statement>> {
        match pair.as_rule() {
            Rule::list => self.walk_list(pair),
            Rule::brace_group => self.walk_group(pair),
            _ => Ok(Vec::new()),
        }
    }

    fn walk_for_in(&mut self, pair: Pair<Rule>) -> R<Statement> {
        let mut var = String::new();
        let mut iter = None;
        let mut static_iter: Option<Vec<String>> = None;
        let mut failglob_checks: Vec<(String, Expression)> = Vec::new();
        let mut body_source: Option<(Rule, String)> = None;
        let mut body = Vec::new();
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::name => var = inner.as_str().to_string(),
                Rule::for_in_clause => {
                    let mut static_words = Vec::new();
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::word_list {
                            let word_pairs = p.clone().into_inner().collect::<Vec<_>>();
                            let has_shell_expansion = word_pairs.iter().any(|w| {
                                self.failglob_pattern_text(w)
                                    .is_some_and(|text| word_text_has_glob_meta(&text))
                            });
                            if self.failglob_enabled && !self.shell_flags.contains(&'f') {
                                for w in &word_pairs {
                                    if let Some(text) = self.failglob_pattern_text(w)
                                        && word_text_has_glob_meta(&text)
                                        && let Some(matches) = self.path_glob_matches_expr(&text)
                                    {
                                        failglob_checks.push((text, matches));
                                    }
                                }
                            }
                            let shell_args = self.words_shell_args(word_pairs)?;
                            let mut all_static = true;
                            for arg in &shell_args {
                                if arg.spread {
                                    all_static = false;
                                    break;
                                }
                                if let Some(text) = literal_string(&arg.value) {
                                    static_words.push(text.to_string());
                                } else {
                                    all_static = false;
                                    break;
                                }
                            }
                            iter = Some(shell_args_array(shell_args));
                            if all_static && !has_shell_expansion {
                                static_iter = Some(static_words.clone());
                            }
                        }
                    }
                }
                Rule::list | Rule::brace_group => {
                    body_source = Some((inner.as_rule(), inner.as_str().to_string()));
                    let saved_variable_values = self.variable_values.clone();
                    let saved_array_values = self.array_values.clone();
                    let saved_indexed_arrays = self.indexed_arrays.clone();
                    let saved_assoc_arrays = self.assoc_arrays.clone();
                    body = self.loop_body(inner)?;
                    self.variable_values = saved_variable_values;
                    self.array_values = saved_array_values;
                    self.indexed_arrays = saved_indexed_arrays;
                    self.assoc_arrays = saved_assoc_arrays;
                }
                _ => {}
            }
        }
        if let (Some(values), Some((rule, source))) = (static_iter, body_source.as_ref()) {
            if !source.contains("return")
                && (source.contains("${!")
                    || source.contains("printf -v")
                    || source.contains("$(("))
            {
                let saved = self.variable_values.get(&var).cloned();
                self.record_variable(&var);
                let mut out = Vec::new();
                for value in values {
                    out.push(assign_stmt(ident(&var), lit(&value)));
                    self.variable_values.insert(var.clone(), value);
                    let pairs = WashmParser::parse(*rule, source)
                        .map_err(|e| format!("static for body: {e}"))?;
                    for p in pairs {
                        match p.as_rule() {
                            Rule::list => out.extend(self.walk_list(p)?),
                            Rule::brace_group => out.extend(self.walk_group(p)?),
                            _ => {}
                        }
                    }
                }
                if let Some(value) = saved {
                    self.variable_values.insert(var, value);
                } else {
                    self.variable_values.remove(&var);
                }
                return Ok(Statement::new(StmtKind::Block(out)));
            }
        }
        self.record_variable(&var);
        let iter_var = self.fresh("__bash_for_item");
        let mut loop_body = Vec::with_capacity(body.len() + 1);
        loop_body.push(assign_stmt(ident(&var), ident(&iter_var)));
        loop_body.extend(body);
        let mut stmt = Statement::new(StmtKind::ForIn {
            var: iter_var,
            key: None,
            iter: iter.unwrap_or_else(|| ident(BASH_ARGS)),
            body: loop_body,
            of: true,
            else_body: None,
            is_async: false,
        });
        for (text, matches) in failglob_checks.into_iter().rev() {
            stmt = Statement::new(StmtKind::If {
                cond: binary(BinOp::StrictEq, member(matches, "length"), int(0)),
                then_body: vec![
                    bash_stderr_stmt(call_named(
                        "__bash_sprintf",
                        vec![lit("bash: no match: %s\n"), lit(&text)],
                    )),
                    assign_stmt(ident("__bash_status"), int(1)),
                ],
                elifs: Vec::new(),
                else_body: Some(vec![stmt]),
            });
        }
        Ok(stmt)
    }

    fn walk_for_arith(&mut self, pair: Pair<Rule>) -> R<Statement> {
        let mut init = None;
        let mut cond = None;
        let mut update = None;
        let mut body = Vec::new();
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::for_arith_init => {
                    let saved_dynamic_arith = self.dynamic_arith;
                    self.dynamic_arith = true;
                    let expr = self.arith_first(inner);
                    self.dynamic_arith = saved_dynamic_arith;
                    init = Some(Box::new(expr_stmt(expr?)));
                }
                Rule::for_arith_cond => {
                    let saved_dynamic_arith = self.dynamic_arith;
                    self.dynamic_arith = true;
                    let expr = self.arith_first(inner);
                    self.dynamic_arith = saved_dynamic_arith;
                    cond = Some(arith_bool(expr?));
                }
                Rule::for_arith_update => {
                    let saved_dynamic_arith = self.dynamic_arith;
                    self.dynamic_arith = true;
                    let expr = self.arith_first(inner);
                    self.dynamic_arith = saved_dynamic_arith;
                    update = Some(expr?);
                }
                Rule::list | Rule::brace_group => {
                    let saved_dynamic_arith = self.dynamic_arith;
                    self.dynamic_arith = true;
                    let walked = self.loop_body(inner);
                    self.dynamic_arith = saved_dynamic_arith;
                    body = walked?;
                }
                _ => {}
            }
        }
        Ok(Statement::new(StmtKind::For {
            init,
            cond,
            update,
            body,
        }))
    }

    fn walk_select(&mut self, pair: Pair<Rule>) -> R<Statement> {
        // `select` shows a menu and reads a reply; the iteration over the
        // choices is kept, the interaction is not.
        let mut var = String::new();
        let mut iter = None;
        let mut body = Vec::new();
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::name => var = inner.as_str().to_string(),
                Rule::select_in_clause => {
                    let mut words = Vec::new();
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::word_list {
                            words.extend(self.words_shell_args(p.into_inner().collect())?);
                        }
                    }
                    iter = Some(shell_args_array(words));
                }
                Rule::list | Rule::brace_group => body = self.loop_body(inner)?,
                _ => {}
            }
        }
        Ok(Statement::new(StmtKind::ForIn {
            var,
            key: None,
            iter: call_named(
                "__bash_select",
                vec![iter.unwrap_or_else(|| ident(BASH_ARGS))],
            ),
            body,
            of: true,
            else_body: None,
            is_async: false,
        }))
    }

    /// `case word in pattern) list ;; … esac` → if/else-if chains.
    ///
    /// `;;` ends the statement, `;&` runs the next body too, `;;&` keeps
    /// testing the remaining patterns (so it starts a new `if`).
    fn walk_case(&mut self, pair: Pair<Rule>) -> R<Vec<Statement>> {
        let tmp = self.fresh("__case");
        let mut items: Vec<(Vec<Pair<Rule>>, Vec<Statement>, String)> = Vec::new();
        let mut subject = None;
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::word if subject.is_none() => subject = Some(self.word_expr(inner)?),
                Rule::case_item => {
                    let mut patterns = Vec::new();
                    let mut body = Vec::new();
                    let mut term = ";;".to_string();
                    for p in inner.into_inner() {
                        match p.as_rule() {
                            Rule::pattern_list => {
                                for pat in p.into_inner() {
                                    if pat.as_rule() == Rule::pattern {
                                        if let Some(w) = pat.into_inner().next() {
                                            patterns.push(w);
                                        }
                                    }
                                }
                            }
                            Rule::list => body = self.walk_list(p)?,
                            Rule::case_terminator => term = p.as_str().to_string(),
                            _ => {}
                        }
                    }
                    items.push((patterns, body, term));
                }
                _ => {}
            }
        }
        let mut out = vec![let_stmt(&tmp, subject.ok_or("case without a word")?)];

        // `;&` fallthrough: append the following bodies until a `;;`.
        let n = items.len();
        let mut bodies: Vec<Vec<Statement>> = Vec::with_capacity(n);
        for i in 0..n {
            let mut b = items[i].1.clone();
            let mut j = i;
            while j < n && items[j].2 == ";&" && j + 1 < n {
                j += 1;
                b.extend(items[j].1.clone());
            }
            bodies.push(b);
        }

        let mut chain: Vec<(Expression, Vec<Statement>)> = Vec::new();
        let flush = |chain: &mut Vec<(Expression, Vec<Statement>)>, out: &mut Vec<Statement>| {
            if chain.is_empty() {
                return;
            }
            let mut it = chain.drain(..);
            let (cond, then_body) = it.next().unwrap();
            let elifs: Vec<(Expression, Vec<Statement>)> = it.collect();
            out.push(Statement::new(StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body: None,
            }));
        };
        for (i, (patterns, _, term)) in items.into_iter().enumerate() {
            let mut cond: Option<Expression> = None;
            for pat in patterns {
                let m = self.pattern_match_expr(ident(&tmp), pat, false)?;
                cond = Some(match cond {
                    None => m,
                    Some(c) => binary(BinOp::Or, c, m),
                });
            }
            let cond = cond.unwrap_or_else(|| Expression::bool(false));
            chain.push((cond, bodies[i].clone()));
            if term == ";;&" {
                flush(&mut chain, &mut out);
            }
        }
        flush(&mut chain, &mut out);
        Ok(out)
    }

    // ── tests ────────────────────────────────────────────────────────────

    fn walk_test_bracket(&mut self, pair: Pair<Rule>) -> R<Expression> {
        let double = pair.as_rule() == Rule::test_double_bracket;
        for inner in pair.into_inner() {
            if inner.as_rule() == Rule::test_expr {
                return self.test_expr(inner, double);
            }
        }
        Err("empty test".into())
    }

    fn single_bracket_runtime_argv(&mut self, pair: &Pair<Rule>) -> R<Option<Expression>> {
        let raw = pair.as_str().trim();
        let Some(inner) = raw.strip_prefix('[').and_then(|s| s.strip_suffix(']')) else {
            return Ok(None);
        };
        let Some(tokens) = shell_word_sources(inner) else {
            return Ok(None);
        };
        let mut needs_runtime = false;
        let mut args = Vec::new();
        for token in tokens {
            let Ok(mut pairs) = WashmParser::parse(Rule::word, &token) else {
                return Ok(None);
            };
            let Some(word) = pairs.next() else {
                continue;
            };
            needs_runtime |=
                word_may_expand_to_multiple_args(&word) || word_is_unquoted_expansion_only(&word);
            args.extend(self.words_shell_args(vec![word])?);
        }
        if needs_runtime {
            Ok(Some(shell_args_array(args)))
        } else {
            Ok(None)
        }
    }

    fn single_bracket_one_arg_expr(&mut self, pair: &Pair<Rule>) -> R<Option<Expression>> {
        let raw = pair.as_str().trim();
        let Some(inner) = raw.strip_prefix('[').and_then(|s| s.strip_suffix(']')) else {
            return Ok(None);
        };
        let Some(tokens) = shell_word_sources(inner) else {
            return Ok(None);
        };
        if tokens.len() != 1 {
            return Ok(None);
        }
        let mut pairs = WashmParser::parse(Rule::word, &tokens[0])
            .map_err(|e| format!("single bracket word: {e}"))?;
        let Some(word) = pairs.next() else {
            return Ok(Some(Expression::bool(false)));
        };
        Ok(Some(binary(BinOp::NotEq, self.word_expr(word)?, lit(""))))
    }

    fn test_expr(&mut self, pair: Pair<Rule>, double: bool) -> R<Expression> {
        for inner in pair.into_inner() {
            if inner.as_rule() == Rule::test_and_or {
                return self.test_and_or(inner, double);
            }
        }
        Err("empty test expression".into())
    }

    fn test_and_or(&mut self, pair: Pair<Rule>, double: bool) -> R<Expression> {
        // `&&` binds tighter than `||` inside [[ ]]; -a tighter than -o in [ ].
        let mut or_terms: Vec<Expression> = Vec::new();
        let mut current: Option<Expression> = None;
        let mut pending_or = false;
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::test_and_op => {}
                Rule::test_or_op => pending_or = true,
                Rule::test_not => {
                    let e = self.test_not(inner, double)?;
                    if pending_or {
                        if let Some(c) = current.take() {
                            or_terms.push(c);
                        }
                        current = Some(e);
                        pending_or = false;
                    } else {
                        current = Some(match current.take() {
                            None => e,
                            Some(c) => binary(BinOp::And, c, e),
                        });
                    }
                }
                _ => {}
            }
        }
        if let Some(c) = current {
            or_terms.push(c);
        }
        let mut it = or_terms.into_iter();
        let mut acc = it.next().ok_or("empty test")?;
        for t in it {
            acc = binary(BinOp::Or, acc, t);
        }
        Ok(acc)
    }

    fn test_not(&mut self, pair: Pair<Rule>, double: bool) -> R<Expression> {
        let mut negate = false;
        let mut result = None;
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::test_bang => negate = !negate,
                Rule::test_paren => {
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::test_expr {
                            result = Some(self.test_expr(p, double)?);
                        }
                    }
                }
                Rule::test_unary_file => {
                    let mut op = String::new();
                    let mut w = None;
                    for p in inner.into_inner() {
                        match p.as_rule() {
                            Rule::test_unary_file_op => op = p.as_str().to_string(),
                            Rule::word => w = Some(self.word_expr(p)?),
                            _ => {}
                        }
                    }
                    let w = w.ok_or("file test without operand")?;
                    result = Some(match op.as_str() {
                        "-e" | "-a" => binary(
                            BinOp::And,
                            binary(BinOp::StrictNotEq, w.clone(), lit("")),
                            call_named("__bash_file_exists", vec![bash_path_expr(w)]),
                        ),
                        "-f" => binary(
                            BinOp::And,
                            binary(BinOp::StrictNotEq, w.clone(), lit("")),
                            call_named("__bash_is_file", vec![bash_path_expr(w)]),
                        ),
                        "-d" => binary(
                            BinOp::And,
                            binary(BinOp::StrictNotEq, w.clone(), lit("")),
                            call_named("__bash_is_dir", vec![bash_path_expr(w)]),
                        ),
                        "-s" => binary(
                            BinOp::And,
                            binary(BinOp::StrictNotEq, w.clone(), lit("")),
                            binary(
                                BinOp::Gt,
                                call_named("__bash_file_size", vec![bash_path_expr(w)]),
                                int(0),
                            ),
                        ),
                        "-o" => shell_option_test_expr(&w, &self.shell_flags),
                        "-t" => bash_spawn_test_expr(vec![lit(&op), w]),
                        _ => bash_spawn_test_expr(vec![lit(&op), bash_path_expr(w)]),
                    });
                }
                Rule::test_unary_string => {
                    let mut op = String::new();
                    let mut w = None;
                    for p in inner.into_inner() {
                        match p.as_rule() {
                            Rule::test_unary_string_op => op = p.as_str().to_string(),
                            Rule::word => w = Some(p),
                            _ => {}
                        }
                    }
                    let w = w.ok_or("string test without operand")?;
                    result = Some(match op.as_str() {
                        "-z" => binary(BinOp::Eq, self.word_expr(w)?, lit("")),
                        "-n" => binary(BinOp::NotEq, self.word_expr(w)?, lit("")),
                        "-v" => {
                            let name = self.static_word_text(&w).unwrap_or_default();
                            self.variable_exists_expr(&name)?
                        }
                        "-R" => {
                            let name = self.static_word_text(&w).unwrap_or_default();
                            Expression::bool(self.namerefs.contains_key(&name))
                        }
                        _ => shell_option_test_expr(&self.word_expr(w)?, &self.shell_flags),
                    });
                }
                Rule::test_binary_file => {
                    let mut op = String::new();
                    let mut ws = Vec::new();
                    for p in inner.into_inner() {
                        match p.as_rule() {
                            Rule::test_binary_file_op => op = p.as_str().to_string(),
                            Rule::word => ws.push(self.word_expr(p)?),
                            _ => {}
                        }
                    }
                    let mut args = vec![lit(&op)];
                    result = Some(match op.as_str() {
                        "-ef" | "-nt" | "-ot" if ws.len() == 2 => bash_spawn_test_expr(vec![
                            bash_path_expr(ws[0].clone()),
                            lit(&op),
                            bash_path_expr(ws[1].clone()),
                        ]),
                        _ => {
                            args.extend(ws);
                            bash_spawn_test_expr(args)
                        }
                    });
                }
                Rule::test_binary_arith => {
                    let mut op = String::new();
                    let mut ws = Vec::new();
                    for p in inner.into_inner() {
                        match p.as_rule() {
                            Rule::test_binary_arith_op => op = p.as_str().to_string(),
                            Rule::word => ws.push(p),
                            _ => {}
                        }
                    }
                    if ws.len() != 2 {
                        return Err("arithmetic test needs two operands".into());
                    }
                    let right = ws.pop().unwrap();
                    let left = ws.pop().unwrap();
                    // [[ ]] evaluates operands as arithmetic expressions; [ ] needs integers.
                    let (l, r) = if double {
                        (
                            self.arith_operand_from_word(left)?,
                            self.arith_operand_from_word(right)?,
                        )
                    } else {
                        (
                            to_number(self.word_expr(left)?),
                            to_number(self.word_expr(right)?),
                        )
                    };
                    let bop = match op.as_str() {
                        "-eq" => BinOp::Eq,
                        "-ne" => BinOp::NotEq,
                        "-lt" => BinOp::Lt,
                        "-le" => BinOp::LtEq,
                        "-gt" => BinOp::Gt,
                        _ => BinOp::GtEq,
                    };
                    result = Some(binary(bop, l, r));
                }
                Rule::test_binary_string => {
                    let mut op = String::new();
                    let mut ws = Vec::new();
                    for p in inner.into_inner() {
                        match p.as_rule() {
                            Rule::test_binary_string_op => op = p.as_str().to_string(),
                            Rule::word => ws.push(p),
                            _ => {}
                        }
                    }
                    if ws.len() != 2 {
                        return Err("string test needs two operands".into());
                    }
                    let right = ws.pop().unwrap();
                    let left = self.word_expr(ws.pop().unwrap())?;
                    result = Some(match op.as_str() {
                        "=" | "==" => self.pattern_match_expr(left, right, !double)?,
                        "!=" => unary(UnaryOp::Not, self.pattern_match_expr(left, right, !double)?),
                        "<" => binary(BinOp::Lt, left, self.word_expr(right)?),
                        _ => binary(BinOp::Gt, left, self.word_expr(right)?),
                    });
                }
                Rule::test_regex_match => {
                    let mut subject = None;
                    let mut re = None;
                    for p in inner.into_inner() {
                        match p.as_rule() {
                            Rule::word => subject = Some(self.word_expr(p)?),
                            Rule::regex_word => re = Some(self.regex_word_expr(p)?),
                            _ => {}
                        }
                    }
                    let subject = subject.ok_or("regex test without subject")?;
                    let re = re.ok_or("regex test without pattern")?;
                    let re = if self.nocasematch_enabled {
                        regexp(re, "i")
                    } else {
                        re
                    };
                    // BASH_REMATCH holds the match and its groups ([] on failure).
                    result = Some(sequence(vec![
                        assign_expr(
                            ident("BASH_REMATCH"),
                            Expression::new(ExprKind::NullCoalesce {
                                left: Box::new(regex_exec_expr(re, subject)),
                                right: Box::new(array(Vec::new())),
                            }),
                        ),
                        binary(BinOp::Gt, member(ident("BASH_REMATCH"), "length"), int(0)),
                    ]));
                }
                Rule::word => {
                    // A lone word is true when non-empty.
                    result = Some(binary(BinOp::NotEq, self.word_expr(inner)?, lit("")));
                }
                _ => {}
            }
        }
        let e = result.ok_or("empty test term")?;
        Ok(if negate { unary(UnaryOp::Not, e) } else { e })
    }

    /// `subject == pattern`: literal patterns compare as strings; glob
    /// patterns are normalized to ECMA RegExp tests.
    fn pattern_match_expr(
        &mut self,
        subject: Expression,
        pattern: Pair<Rule>,
        literal_only: bool,
    ) -> R<Expression> {
        let flags = if self.nocasematch_enabled { "i" } else { "" };
        if !literal_only {
            if let Some(text) = self.static_unquoted_param_pattern_text(&pattern) {
                if text == "*" {
                    return Ok(Expression::bool(true));
                }
                if word_text_has_glob_meta(&text) {
                    if let Some(re) = glob_to_regex(&text, false) {
                        return Ok(regex_test_expr(
                            regexp(lit(&format!("^(?:{re})$")), flags),
                            subject,
                        ));
                    }
                }
                return Ok(if self.nocasematch_enabled {
                    binary(
                        BinOp::Eq,
                        method(subject, "toLowerCase", vec![]),
                        lit(&text.to_lowercase()),
                    )
                } else {
                    binary(BinOp::Eq, subject, lit(&text))
                });
            }
        }
        let is_glob = !literal_only && word_has_unquoted_glob(&pattern);
        if !is_glob {
            let pat = self.word_expr(pattern)?;
            return Ok(if self.nocasematch_enabled && !literal_only {
                binary(
                    BinOp::Eq,
                    method(subject, "toLowerCase", vec![]),
                    method(pat, "toLowerCase", vec![]),
                )
            } else {
                binary(BinOp::Eq, subject, pat)
            });
        }
        if let Some(text) = glob_word_text(&pattern) {
            if text == "*" {
                return Ok(Expression::bool(true));
            }
            if let Some(re) = glob_to_regex(&text, false) {
                return Ok(regex_test_expr(
                    regexp(lit(&format!("^(?:{re})$")), flags),
                    subject,
                ));
            }
        }
        let pat = self.word_expr(pattern)?;
        Ok(regex_test_expr(regexp(pat, flags), subject))
    }

    fn static_unquoted_param_pattern_text(&self, pair: &Pair<Rule>) -> Option<String> {
        let name = whole_unquoted_param_name(pair.as_str().trim())?;
        let resolved = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        self.variable_values.get(&resolved).cloned()
    }

    fn static_glob_pattern_text(&self, pair: &Pair<Rule>) -> Option<String> {
        let mut out = String::new();
        for part in pair.clone().into_inner() {
            self.push_static_glob_pattern_part(&part, false, &mut out)?;
        }
        Some(out)
    }

    fn static_replacement_text(&self, pair: &Pair<Rule>) -> Option<String> {
        let mut out = String::new();
        for part in pair.clone().into_inner() {
            match part.as_rule() {
                Rule::brace_text | Rule::pattern_text | Rule::bare_word => {
                    out.push_str(part.as_str())
                }
                Rule::single_quoted_string => {
                    let s = part.as_str();
                    out.push_str(&s[1..s.len() - 1]);
                }
                Rule::ansi_c_quoting => {
                    let s = part.as_str();
                    out.push_str(&decode_ansi_c(&s[2..s.len() - 1]));
                }
                Rule::quoted_string | Rule::locale_quoting => {
                    out.push_str(&self.static_part_text(&part)?);
                }
                Rule::simple_param | Rule::braced_param | Rule::arithmetic_expansion => {
                    out.push_str(&self.static_part_text(&part)?);
                }
                Rule::dollar_paren_subst => return None,
                _ => return None,
            }
        }
        Some(out)
    }

    fn push_static_glob_pattern_part(
        &self,
        pair: &Pair<Rule>,
        quoted: bool,
        out: &mut String,
    ) -> Option<()> {
        match pair.as_rule() {
            Rule::brace_text | Rule::pattern_text => {
                if quoted {
                    out.push_str(&glob_escape(pair.as_str()));
                } else {
                    out.push_str(pair.as_str());
                }
            }
            Rule::bare_word => {
                let text = unescape_bare(pair.as_str());
                if quoted {
                    out.push_str(&glob_escape(&text));
                } else {
                    out.push_str(&text);
                }
            }
            Rule::single_quoted_string => {
                let s = pair.as_str();
                out.push_str(&glob_escape(&s[1..s.len() - 1]));
            }
            Rule::ansi_c_quoting => {
                let s = pair.as_str();
                out.push_str(&glob_escape(&decode_ansi_c(&s[2..s.len() - 1])));
            }
            Rule::quoted_string | Rule::locale_quoting => {
                for inner in pair.clone().into_inner() {
                    match inner.as_rule() {
                        Rule::dq_text => out.push_str(&glob_escape(inner.as_str())),
                        Rule::dq_escape => out.push_str(&glob_escape(&inner.as_str()[1..])),
                        _ => {
                            let text = self.static_part_text(&inner)?;
                            out.push_str(&glob_escape(&text));
                        }
                    }
                }
            }
            Rule::simple_param | Rule::braced_param | Rule::arithmetic_expansion => {
                let text = self.static_part_text(pair)?;
                if quoted {
                    out.push_str(&glob_escape(&text));
                } else {
                    out.push_str(&text);
                }
            }
            Rule::dollar_paren_subst => return None,
            _ => return None,
        }
        Some(())
    }

    fn regex_word_expr(&mut self, pair: Pair<Rule>) -> R<Expression> {
        let mut parts = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::regex_text => parts.push(Part::Text(p.as_str().to_string())),
                // Quoted regex text is literal: escape its metacharacters.
                Rule::quoted_string | Rule::single_quoted_string => {
                    let inner = self.part_to_parts(p)?;
                    for part in inner {
                        match part {
                            Part::Text(t) => parts.push(Part::Text(regex_escape(&t))),
                            Part::Expr(e) => {
                                parts.push(Part::Expr(call_named("__bash_regex_quote", vec![e])))
                            }
                        }
                    }
                }
                _ => parts.push(Part::Expr(self.part_expr(p)?)),
            }
        }
        Ok(parts_to_expr(parts))
    }

    // ── arithmetic ───────────────────────────────────────────────────────

    fn arith_cmd_value(&mut self, pair: Pair<Rule>) -> R<Expression> {
        for inner in pair.into_inner() {
            if inner.as_rule() == Rule::arithmetic_expr {
                return self.arith_expr(inner);
            }
        }
        Ok(int(0))
    }

    fn arith_first(&mut self, pair: Pair<Rule>) -> R<Expression> {
        for inner in pair.into_inner() {
            if inner.as_rule() == Rule::arithmetic_expr {
                return self.arith_expr(inner);
            }
        }
        Err("empty arithmetic expression".into())
    }

    fn parse_arith_text(&mut self, text: &str) -> R<Expression> {
        let expanded;
        let text = if text.contains('$') {
            expanded = self
                .static_arith_parameter_substitution(text)
                .unwrap_or_else(|| text.to_string());
            expanded.as_str()
        } else {
            text
        };
        let normalized;
        let text = if text.contains('\n') || text.contains('\r') {
            normalized = text.replace(['\r', '\n'], " ");
            normalized.as_str()
        } else {
            text
        };
        let text = text.trim();
        let pairs = WashmParser::parse(Rule::arithmetic_expr, text)
            .map_err(|e| format!("arithmetic: {e}"))?;
        for p in pairs {
            if p.as_rule() == Rule::arithmetic_expr {
                if p.as_str().len() != text.trim().len() {
                    return Err(format!("arithmetic: trailing input in `{text}`"));
                }
                return self.arith_expr(p);
            }
        }
        Ok(int(0))
    }

    /// An operand of `[[ x -eq y ]]`: the word's text is an arithmetic expression.
    fn arith_operand_from_word(&mut self, w: Pair<Rule>) -> R<Expression> {
        if word_is_plain(&w) {
            let text = word_source_text(&w);
            if let Ok(e) = self.parse_arith_text(&text) {
                return Ok(e);
            }
        }
        Ok(to_number(self.word_expr(w)?))
    }

    fn arith_expr(&mut self, pair: Pair<Rule>) -> R<Expression> {
        let inner = pair.into_inner().next().ok_or("empty arithmetic")?;
        self.arith_node(inner)
    }

    fn arith_node(&mut self, pair: Pair<Rule>) -> R<Expression> {
        match pair.as_rule() {
            Rule::arithmetic_expr => self.arith_expr(pair),
            Rule::arithmetic_comma => {
                let items: Vec<Pair<Rule>> = pair
                    .into_inner()
                    .filter(|p| p.as_rule() != Rule::arith_comma_op)
                    .collect();
                if items.len() == 1 {
                    return self.arith_node(items.into_iter().next().unwrap());
                }
                let mut seq = Vec::new();
                for i in items {
                    seq.push(self.arith_node(i)?);
                }
                Ok(sequence(seq))
            }
            Rule::arithmetic_assign => {
                let items: Vec<Pair<Rule>> = pair.into_inner().collect();
                if items.len() == 1 {
                    return self.arith_node(items.into_iter().next().unwrap());
                }
                let mut it = items.into_iter();
                let target_pair = it.next().unwrap();
                if self.arith_lvalue_is_readonly(&target_pair) {
                    return Ok(int(0));
                }
                self.forget_static_arith_lvalue(&target_pair);
                let target = self.arith_variable_ref(target_pair)?;
                let op = it
                    .next()
                    .map(|p| p.as_str().to_string())
                    .unwrap_or_else(|| "=".into());
                let value = self.arith_node(it.next().ok_or("assignment without value")?)?;
                let value = match op.as_str() {
                    "=" => value,
                    _ => {
                        let bop = match op.as_str() {
                            "+=" => "__bash_i64_add",
                            "-=" => "__bash_i64_sub",
                            "*=" => "__bash_i64_mul",
                            "/=" => "__bash_i64_div",
                            "%=" => "__bash_i64_rem",
                            "**=" => "__bash_i64_pow",
                            "<<=" => "__bash_i64_shl",
                            ">>=" => "__bash_i64_shr",
                            "&=" => "__bash_i64_and",
                            "|=" => "__bash_i64_or",
                            _ => "__bash_i64_xor",
                        };
                        arith_bin(bop, target.clone(), value)
                    }
                };
                Ok(assign_expr(target, arith_string(value)))
            }
            Rule::arithmetic_ternary => {
                let items: Vec<Pair<Rule>> = pair
                    .into_inner()
                    .filter(|p| !matches!(p.as_rule(), Rule::arith_question | Rule::arith_colon))
                    .collect();
                if items.len() == 1 {
                    return self.arith_node(items.into_iter().next().unwrap());
                }
                let mut it = items.into_iter();
                let c = self.arith_node(it.next().unwrap())?;
                let t = self.arith_node(it.next().ok_or("ternary without then")?)?;
                let e = self.arith_node(it.next().ok_or("ternary without else")?)?;
                Ok(ternary(binary(BinOp::NotEq, c, int(0)), t, e))
            }
            Rule::arithmetic_or | Rule::arithmetic_and => {
                let is_or = pair.as_rule() == Rule::arithmetic_or;
                let items: Vec<Pair<Rule>> = pair
                    .into_inner()
                    .filter(|p| !matches!(p.as_rule(), Rule::arith_or_op | Rule::arith_and_op))
                    .collect();
                if items.len() == 1 {
                    return self.arith_node(items.into_iter().next().unwrap());
                }
                let mut acc: Option<Expression> = None;
                for i in items {
                    let v = arith_bool(self.arith_node(i)?);
                    acc = Some(match acc {
                        None => v,
                        Some(a) => binary(if is_or { BinOp::Or } else { BinOp::And }, a, v),
                    });
                }
                Ok(arith_bool_to_i64(acc.unwrap()))
            }
            Rule::arithmetic_bitor
            | Rule::arithmetic_bitxor
            | Rule::arithmetic_bitand
            | Rule::arithmetic_shift
            | Rule::arithmetic_add
            | Rule::arithmetic_mul => self.arith_left_assoc(pair),
            Rule::arithmetic_eq | Rule::arithmetic_cmp => {
                let items: Vec<Pair<Rule>> = pair.into_inner().collect();
                if items.len() == 1 {
                    return self.arith_node(items.into_iter().next().unwrap());
                }
                let mut it = items.into_iter();
                let mut acc = self.arith_node(it.next().unwrap())?;
                while let Some(op) = it.next() {
                    let rhs =
                        self.arith_node(it.next().ok_or("comparison without right operand")?)?;
                    let bop = match op.as_str() {
                        "==" => "__bash_i64_eq",
                        "!=" => "__bash_i64_ne",
                        "<" => "__bash_i64_lt",
                        ">" => "__bash_i64_gt",
                        "<=" => "__bash_i64_le",
                        _ => "__bash_i64_ge",
                    };
                    acc = arith_cmp_i64(bop, acc, rhs);
                }
                Ok(acc)
            }
            Rule::arithmetic_pow => {
                // Right-associative.
                let items: Vec<Pair<Rule>> = pair
                    .into_inner()
                    .filter(|p| p.as_rule() != Rule::arith_pow_op)
                    .collect();
                let mut exprs = Vec::new();
                for i in items {
                    exprs.push(self.arith_node(i)?);
                }
                let mut acc = exprs.pop().ok_or("empty power")?;
                while let Some(base) = exprs.pop() {
                    acc = arith_bin("__bash_i64_pow", base, acc);
                }
                Ok(acc)
            }
            Rule::arithmetic_unary => {
                let items: Vec<Pair<Rule>> = pair.into_inner().collect();
                if items.len() == 1 {
                    return self.arith_node(items.into_iter().next().unwrap());
                }
                let mut it = items.into_iter();
                let op = it.next().unwrap().as_str().to_string();
                let operand_pair = it.next().ok_or("unary without operand")?;
                match op.as_str() {
                    "++" | "--" => {
                        if self.arith_lvalue_is_readonly(&operand_pair) {
                            return Ok(int(0));
                        }
                        self.forget_static_arith_lvalue(&operand_pair);
                        let target = self.arith_lvalue_from_unary(operand_pair)?;
                        Ok(self.arith_update_expr(target, if op == "++" { 1 } else { -1 }, true))
                    }
                    _ => {
                        let v = self.arith_node(operand_pair)?;
                        Ok(match op.as_str() {
                            "-" => arith_unary("__bash_i64_neg", v),
                            "+" => arith_i64(v),
                            "!" => arith_bool_to_i64(unary(UnaryOp::Not, arith_bool(v))),
                            _ => arith_unary("__bash_i64_not", v),
                        })
                    }
                }
            }
            Rule::arithmetic_postfix => {
                let items: Vec<Pair<Rule>> = pair.into_inner().collect();
                if items.len() == 1 {
                    return self.arith_primary(items.into_iter().next().unwrap());
                }
                let mut it = items.into_iter();
                let target_pair = it.next().unwrap();
                if self.arith_lvalue_is_readonly(&target_pair) {
                    return Ok(int(0));
                }
                self.forget_static_arith_lvalue(&target_pair);
                let target = self.arith_primary_lvalue(target_pair)?;
                let op = it.next().unwrap().as_str().to_string();
                Ok(self.arith_update_expr(target, if op == "++" { 1 } else { -1 }, false))
            }
            _ => self.arith_primary(pair),
        }
    }

    fn arith_left_assoc(&mut self, pair: Pair<Rule>) -> R<Expression> {
        let items: Vec<Pair<Rule>> = pair.into_inner().collect();
        if items.len() == 1 {
            return self.arith_node(items.into_iter().next().unwrap());
        }
        let mut it = items.into_iter();
        let mut acc = self.arith_node(it.next().unwrap())?;
        while let Some(op) = it.next() {
            let rhs = self.arith_node(it.next().ok_or("binary operator without right operand")?)?;
            let bop = match op.as_str() {
                "+" => "__bash_i64_add",
                "-" => "__bash_i64_sub",
                "*" => "__bash_i64_mul",
                "/" => "__bash_i64_div",
                "%" => "__bash_i64_rem",
                "<<" => "__bash_i64_shl",
                ">>" => "__bash_i64_shr",
                "&" => "__bash_i64_and",
                "|" => "__bash_i64_or",
                _ => "__bash_i64_xor",
            };
            acc = arith_bin(bop, acc, rhs);
        }
        Ok(acc)
    }

    fn arith_primary(&mut self, pair: Pair<Rule>) -> R<Expression> {
        match pair.as_rule() {
            Rule::arith_paren => {
                self.arith_expr(pair.into_inner().next().ok_or("empty parentheses")?)
            }
            Rule::arith_number => parse_bash_integer(pair.as_str())
                .map(bigint)
                .ok_or_else(|| format!("invalid arithmetic literal {}", pair.as_str())),
            Rule::arith_variable => self.arith_variable_value(pair),
            Rule::arithmetic_expansion => Ok(call_named(
                "__bash_string",
                vec![self.arith_expansion_expr(pair)?],
            )),
            Rule::simple_param
            | Rule::braced_param
            | Rule::dollar_paren_subst
            | Rule::quoted_string
            | Rule::single_quoted_string => Ok(to_number(self.part_expr(pair)?)),
            other => Err(format!(
                "unsupported arithmetic primary {other:?}: {}",
                pair.as_str()
            )),
        }
    }

    fn arith_primary_lvalue(&mut self, pair: Pair<Rule>) -> R<Expression> {
        match pair.as_rule() {
            Rule::arith_variable | Rule::arithmetic_lvalue => self.arith_variable_ref(pair),
            Rule::arith_paren => {
                self.arith_primary_lvalue(pair.into_inner().next().ok_or("empty parentheses")?)
            }
            _ => Err(format!("`++`/`--` need a variable, got {}", pair.as_str())),
        }
    }

    fn arith_lvalue_from_unary(&mut self, pair: Pair<Rule>) -> R<Expression> {
        let mut cur = pair;
        loop {
            match cur.as_rule() {
                Rule::arithmetic_unary | Rule::arithmetic_postfix => {
                    let items: Vec<Pair<Rule>> = cur.into_inner().collect();
                    if items.len() != 1 {
                        return Err("`++`/`--` need a plain variable".into());
                    }
                    cur = items.into_iter().next().unwrap();
                }
                _ => return self.arith_primary_lvalue(cur),
            }
        }
    }

    fn arith_variable_value(&mut self, pair: Pair<Rule>) -> R<Expression> {
        if let Some(name) = arith_simple_name(&pair) {
            let resolved = self
                .resolve_nameref_text(&name)
                .unwrap_or_else(|| name.clone());
            if self.arith_stack.iter().any(|active| active == &resolved) {
                return Ok(sequence(vec![
                    assign_expr(ident("__bash_status"), int(1)),
                    int(0),
                ]));
            }
            if !self.dynamic_arith {
                if let Some(value) = self.variable_values.get(&resolved).cloned() {
                    let text = value.trim();
                    if text.is_empty() {
                        return Ok(int(0));
                    }
                    if let Some(n) = parse_bash_integer(text) {
                        return Ok(bigint(n));
                    }
                    self.arith_stack.push(resolved);
                    let parsed = self.parse_arith_text(text);
                    self.arith_stack.pop();
                    return Ok(parsed.unwrap_or_else(|_| {
                        sequence(vec![assign_expr(ident("__bash_status"), int(1)), bigint(0)])
                    }));
                }
            }
        }
        Ok(arith_i64(self.arith_variable_ref(pair)?))
    }

    fn arith_update_expr(&mut self, target: Expression, delta: i64, prefix: bool) -> Expression {
        let saved = self.fresh("__bash_arith_update");
        let current = arith_i64(target.clone());
        let next = arith_bin(
            if delta >= 0 {
                "__bash_i64_add"
            } else {
                "__bash_i64_sub"
            },
            ident(&saved),
            bigint(delta.abs()),
        );
        iife(vec![
            let_stmt(&saved, current),
            expr_stmt(assign_expr(target, arith_string(next.clone()))),
            Statement::new(StmtKind::Return(Some(if prefix {
                next
            } else {
                ident(&saved)
            }))),
        ])
    }

    fn simple_arithmetic_update_expr(&mut self, raw: &str) -> Option<Expression> {
        let inner = arithmetic_cmd_inner(raw)?.trim();
        let (prefix, name, delta) = if let Some(name) = inner.strip_prefix("++") {
            (true, name.trim(), 1)
        } else if let Some(name) = inner.strip_prefix("--") {
            (true, name.trim(), -1)
        } else if let Some(name) = inner.strip_suffix("++") {
            (false, name.trim(), 1)
        } else if let Some(name) = inner.strip_suffix("--") {
            (false, name.trim(), -1)
        } else {
            return None;
        };
        if !is_name(name) || self.target_text_is_readonly(name) {
            return None;
        }
        let resolved = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        self.record_variable(&resolved);
        self.variable_values.remove(&resolved);
        Some(self.arith_update_expr(ident(&resolved), delta, prefix))
    }

    /// `name` or `name[sub]` inside arithmetic — a variable reference (no coercion).
    fn arith_variable_ref(&mut self, pair: Pair<Rule>) -> R<Expression> {
        let mut name = String::new();
        let mut sub = None;
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::name => name = p.as_str().to_string(),
                Rule::arith_subscript => {
                    let inner = p.into_inner().next().ok_or("empty subscript")?;
                    sub = Some(match inner.as_rule() {
                        Rule::subscript_all => Subscript::All,
                        _ => Subscript::Key(self.arith_node(inner)?),
                    });
                }
                _ => {}
            }
        }
        let target = self
            .resolve_nameref_text(&name)
            .unwrap_or_else(|| name.clone());
        if !target.is_empty() && !target.contains('[') {
            self.record_variable(&target);
        }
        Ok(match sub {
            Some(Subscript::Key(_)) if target.contains('[') => {
                self.name_or_element_target(&target)?
            }
            Some(Subscript::Key(k)) => index(ident(&target), arith_index(k)),
            _ if target.contains('[') => self.name_or_element_target(&target)?,
            _ => ident(&target),
        })
    }

    fn arith_lvalue_is_readonly(&self, pair: &Pair<Rule>) -> bool {
        arith_lvalue_text(pair)
            .as_deref()
            .is_some_and(|name| self.target_text_is_readonly(name))
    }

    fn forget_static_arith_lvalue(&mut self, pair: &Pair<Rule>) {
        let Some(text) = arith_lvalue_text(pair) else {
            return;
        };
        let name = text
            .split_once('[')
            .map(|(base, _)| base)
            .unwrap_or(text.as_str());
        let target = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        self.variable_values.remove(&target);
        self.array_values.remove(&target);
    }

    fn arith_expansion_expr(&mut self, pair: Pair<Rule>) -> R<Expression> {
        if let Some(inner) = arithmetic_expansion_inner(pair.as_str()) {
            if self.dynamic_arith {
                if let Ok(expr) = self.parse_arith_text(inner) {
                    return Ok(expr);
                }
                if let Some(message) = bash_arithmetic_error_message(inner) {
                    return Ok(bash_arithmetic_error_expr(message));
                }
            } else {
                let expanded = if inner.contains('$') {
                    self.static_arith_parameter_substitution(inner)
                } else {
                    None
                };
                let text = expanded.as_deref().unwrap_or(inner);
                if let Some(value) = self.eval_static_arith_text(text) {
                    return Ok(bigint(value));
                }
                if let Ok(expr) = self.parse_arith_text(text) {
                    return Ok(expr);
                }
                if let Some(message) = bash_arithmetic_error_message(text) {
                    return Ok(bash_arithmetic_error_expr(message));
                }
                if inner.contains('$') && expanded.is_none() {
                    if let Ok(expr) = self.parse_arith_text(inner) {
                        return Ok(expr);
                    }
                }
            }
            if let Ok(expr) = self.parse_arith_text(inner) {
                return Ok(expr);
            }
            if let Some(message) = bash_arithmetic_error_message(inner) {
                return Ok(bash_arithmetic_error_expr(message));
            }
        }
        for inner in pair.into_inner() {
            if inner.as_rule() == Rule::arithmetic_expr {
                return self.arith_expr(inner);
            }
        }
        Ok(int(0))
    }

    fn static_arith_parameter_substitution(&self, text: &str) -> Option<String> {
        let mut out = String::with_capacity(text.len());
        let chars: Vec<char> = text.chars().collect();
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
            match chars[i + 1] {
                '{' => {
                    let mut j = i + 2;
                    while j < chars.len() && chars[j] != '}' {
                        j += 1;
                    }
                    if j >= chars.len() {
                        return None;
                    }
                    let name: String = chars[i + 2..j].iter().collect();
                    let value = if is_name(&name) {
                        self.static_name_value(&name)?
                    } else {
                        return None;
                    };
                    out.push_str(&value);
                    i = j + 1;
                }
                c if c.is_ascii_alphabetic() || c == '_' => {
                    let mut j = i + 1;
                    while j < chars.len() && (chars[j].is_ascii_alphanumeric() || chars[j] == '_') {
                        j += 1;
                    }
                    let name: String = chars[i + 1..j].iter().collect();
                    out.push_str(&self.static_name_value(&name)?);
                    i = j;
                }
                c if c.is_ascii_digit() => {
                    let pos = c.to_digit(10)? as usize;
                    if let Some(value) = self.positional_values.get(pos.saturating_sub(1)) {
                        out.push_str(value);
                    }
                    i += 2;
                }
                '#' => {
                    out.push_str(&self.positional_values.len().to_string());
                    i += 2;
                }
                '?' | '$' | '!' => {
                    out.push('0');
                    i += 2;
                }
                _ => return None,
            }
        }
        Some(out)
    }

    // ── words ────────────────────────────────────────────────────────────

    fn words_exprs(&mut self, words: Vec<Pair<Rule>>) -> R<Vec<Expression>> {
        let mut out = Vec::new();
        for w in words {
            match w.as_rule() {
                Rule::word => out.extend(self.expand_word(w)?),
                Rule::close_brace_arg => out.push(lit("}")),
                // `cmd a=b`: an assignment-shaped argument is just a word.
                Rule::assignment_word => out.push(self.assignment_word_as_text(w)?),
                _ => {}
            }
        }
        Ok(out)
    }

    fn words_shell_args(&mut self, words: Vec<Pair<Rule>>) -> R<Vec<ShellArg>> {
        let mut out = Vec::new();
        for w in words {
            match w.as_rule() {
                Rule::word => {
                    if let Some(keys) = self.exact_indirect_array_key_values(&w) {
                        out.push(ShellArg {
                            value: array(keys),
                            spread: true,
                        });
                    } else if let Some(value) = self.exact_transformed_array_word(&w)? {
                        out.push(ShellArg {
                            value,
                            spread: true,
                        });
                    } else if let Some(value) = self.exact_array_at_word(&w)? {
                        out.push(ShellArg {
                            value,
                            spread: true,
                        });
                    } else if let Some(value) = self.exact_process_substitution_word(&w)? {
                        out.push(ShellArg {
                            value,
                            spread: false,
                        });
                    } else if !word_has_quoted_part(&w) {
                        if let Some(text) = self.static_word_text(&w) {
                            if word_has_unquoted_expansion(&w) {
                                out.extend(self.static_split_glob_shell_args(&text));
                            } else if word_has_unquoted_glob(&w) {
                                out.extend(self.static_glob_shell_args(&text));
                            } else {
                                out.push(ShellArg {
                                    value: lit(&text),
                                    spread: false,
                                });
                            }
                        } else if let Some(value) = self.exact_scalar_split_word(&w)? {
                            out.push(ShellArg {
                                value,
                                spread: true,
                            });
                        } else if let Some(value) =
                            self.exact_command_substitution_split_word(&w)?
                        {
                            out.push(ShellArg {
                                value,
                                spread: true,
                            });
                        } else if word_is_unquoted_expansion_only(&w) {
                            let value = self.word_expr(w)?;
                            out.push(ShellArg {
                                value: self.split_and_glob_words_expr(value),
                                spread: true,
                            });
                        } else if let Some(value) = self.path_glob_word(&w) {
                            out.push(ShellArg {
                                value,
                                spread: true,
                            });
                        } else {
                            let expanded = self.expand_word(w)?;
                            for value in expanded {
                                out.extend(self.expanded_unquoted_word_shell_args(value));
                            }
                    }
                } else if let Some(args) = self.static_mixed_word_shell_args(&w)? {
                    out.extend(args);
                } else if let Some(value) = self.path_glob_word(&w) {
                    out.push(ShellArg {
                        value,
                        spread: true,
                    });
                } else if let Some(text) = self.static_word_text(&w) {
                    out.push(ShellArg {
                        value: lit(&text),
                            spread: false,
                        });
                    } else if let Some(value) = self.exact_scalar_split_word(&w)? {
                        out.push(ShellArg {
                            value,
                            spread: true,
                        });
                    } else if let Some(value) = self.exact_command_substitution_split_word(&w)? {
                        out.push(ShellArg {
                            value,
                            spread: true,
                        });
                    } else if word_is_unquoted_expansion_only(&w) {
                        out.push(ShellArg {
                            value: split_bash_words(self.word_expr(w)?),
                            spread: true,
                        });
                } else {
                    out.extend(self.expand_word(w)?.into_iter().map(|value| ShellArg {
                        value,
                            spread: false,
                        }));
                    }
                }
                Rule::close_brace_arg => out.push(ShellArg {
                    value: lit("}"),
                    spread: false,
                }),
                Rule::assignment_word => out.push(ShellArg {
                    value: self.assignment_word_as_text(w)?,
                    spread: false,
                }),
                _ => {}
            }
        }
        Ok(out)
    }

    fn exact_process_substitution_word(&mut self, pair: &Pair<Rule>) -> R<Option<Expression>> {
        if pair.as_rule() != Rule::word {
            return Ok(None);
        }
        let raw = pair.as_str().trim();
        if raw.starts_with('"') || raw.starts_with('\'') {
            return Ok(None);
        }
        let mut inner = pair.clone().into_inner();
        let Some(part) = inner.next() else {
            return Ok(None);
        };
        if inner.next().is_some() || part.as_rule() != Rule::process_substitution {
            return Ok(None);
        }
        Ok(Some(self.part_expr(part)?))
    }

    fn exact_command_substitution_split_word(
        &mut self,
        pair: &Pair<Rule>,
    ) -> R<Option<Expression>> {
        if pair.as_rule() != Rule::word {
            return Ok(None);
        }
        let raw = pair.as_str().trim();
        if raw.starts_with('"') || raw.starts_with('\'') {
            return Ok(None);
        }
        let mut inner = pair.clone().into_inner();
        let Some(part) = inner.next() else {
            return Ok(None);
        };
        if inner.next().is_some() {
            return Ok(None);
        }
        match part.as_rule() {
            Rule::dollar_paren_subst | Rule::backtick_subst | Rule::brace_command_subst => {
                let value = self.part_expr(part)?;
                Ok(Some(self.split_and_glob_words_expr(value)))
            }
            _ => Ok(None),
        }
    }

    fn here_string_word_expr(&mut self, word: Pair<Rule>) -> R<Expression> {
        if let Some(joined) = self.exact_array_join_word(&word)? {
            return Ok(joined);
        }
        self.word_expr(word)
    }

    fn exact_array_join_word(&mut self, pair: &Pair<Rule>) -> R<Option<Expression>> {
        let raw = pair.as_str().trim();
        let quoted = raw.starts_with('"') && raw.ends_with('"') && raw.len() >= 2;
        let text = if quoted { &raw[1..raw.len() - 1] } else { raw };
        match text {
            "$@" => return Ok(Some(array_join(ident(BASH_ARGS), lit(" ")))),
            "$*" => return Ok(Some(array_join(ident(BASH_ARGS), ifs_join_sep()))),
            _ => {}
        }
        let Some(inner) = text.strip_prefix("${").and_then(|s| s.strip_suffix('}')) else {
            return Ok(None);
        };
        let (name, sep) = if let Some(name) = inner.strip_suffix("[@]") {
            (name, lit(" "))
        } else if let Some(name) = inner.strip_suffix("[*]") {
            (name, ifs_join_sep())
        } else {
            return Ok(None);
        };
        if !is_name(name) {
            return Ok(None);
        }
        let target = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        if target.contains('[') {
            return Ok(None);
        }
        Ok(Some(array_join(ident(&target), sep)))
    }

    fn exact_array_at_word(&mut self, pair: &Pair<Rule>) -> R<Option<Expression>> {
        let raw = pair.as_str().trim();
        let quoted = raw.starts_with('"') && raw.ends_with('"') && raw.len() >= 2;
        let text = if quoted { &raw[1..raw.len() - 1] } else { raw };
        if !quoted && text == "$*" {
            return Ok(Some(split_bash_words(array_join(
                ident(BASH_ARGS),
                lit(" "),
            ))));
        }
        if text == "$@" {
            return Ok(Some(if quoted {
                ident(BASH_ARGS)
            } else {
                filter_non_empty_array_expr(ident(BASH_ARGS))
            }));
        }
        let Some(inner) = text.strip_prefix("${").and_then(|s| s.strip_suffix('}')) else {
            return Ok(None);
        };
        if let Some(rest) = inner
            .strip_prefix("@:")
            .or_else(|| inner.strip_prefix("*:"))
        {
            if let Some((offset, length)) = parse_positional_slice_text(rest) {
                return Ok(Some(positional_slice_expr(
                    ident(BASH_ARGS),
                    offset,
                    length,
                )));
            }
        }
        let Some(name) = inner.strip_suffix("[@]") else {
            return Ok(None);
        };
        if !is_name(name) {
            return Ok(None);
        }
        let target = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        if target.contains('[') {
            return Ok(None);
        }
        Ok(Some(if quoted {
            ident(&target)
        } else {
            filter_non_empty_array_expr(ident(&target))
        }))
    }

    fn exact_indirect_array_key_values(&self, pair: &Pair<Rule>) -> Option<Vec<Expression>> {
        let raw = pair.as_str().trim();
        let quoted = raw.starts_with('"') && raw.ends_with('"') && raw.len() >= 2;
        let text = if quoted { &raw[1..raw.len() - 1] } else { raw };
        let inner = text.strip_prefix("${!")?.strip_suffix('}')?;
        let name = inner
            .strip_suffix("[@]")
            .or_else(|| inner.strip_suffix("[*]"))?;
        if !is_name(name) {
            return None;
        }
        let target = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        let keys = self.static_array_keys(&target)?;
        Some(keys.into_iter().map(|key| lit(&key)).collect())
    }

    fn exact_scalar_split_word(&mut self, pair: &Pair<Rule>) -> R<Option<Expression>> {
        let raw = pair.as_str().trim();
        if let Some(special) = whole_unquoted_special_param(raw) {
            return Ok(Some(array(vec![special_param_value(special)])));
        }
        let Some(name) = whole_unquoted_param_name(raw) else {
            return Ok(None);
        };
        let target = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        if let Some(value) = self.variable_values.get(&target).cloned() {
            return Ok(Some(shell_args_array(self.static_split_glob_shell_args(
                &value,
            ))));
        }
        let target = self.name_or_element_target(name)?;
        Ok(Some(self.split_and_glob_words_expr(param_value(target))))
    }

    fn static_mixed_word_shell_args(&mut self, pair: &Pair<Rule>) -> R<Option<Vec<ShellArg>>> {
        if pair.as_rule() != Rule::word || !word_has_quoted_part(pair) {
            return Ok(None);
        }
        if let Some(args) = self.static_mixed_word_shell_args_from_raw(pair.as_str()) {
            return Ok(Some(args));
        }
        let mut fields = vec![String::new()];
        let mut saw_unquoted_split = false;
        for part in pair.clone().into_inner() {
            match part.as_rule() {
                Rule::bare_word
                | Rule::assignment_bare_word
                | Rule::brace_seq_text
                | Rule::pattern_text
                | Rule::brace_text => {
                    if let Some(last) = fields.last_mut() {
                        last.push_str(&unescape_bare(part.as_str()));
                    }
                }
                Rule::quoted_string
                | Rule::single_quoted_string
                | Rule::ansi_c_quoting
                | Rule::locale_quoting => {
                    let Some(text) = static_parts_text(&self.part_to_parts(part)?) else {
                        return Ok(None);
                    };
                    if let Some(last) = fields.last_mut() {
                        last.push_str(&text);
                    }
                }
                Rule::simple_param => {
                    let Some(text) = self.static_simple_param_text(&part)? else {
                        return Ok(None);
                    };
                    saw_unquoted_split = true;
                    let split = self.static_split_bash_words(&text);
                    if let Some((first, rest)) = split.split_first() {
                        if let Some(last) = fields.last_mut() {
                            last.push_str(first);
                        }
                        fields.extend(rest.iter().cloned());
                    }
                }
                Rule::braced_param => {
                    let Some(text) = self.static_braced_param_text(&part)? else {
                        return Ok(None);
                    };
                    saw_unquoted_split = true;
                    let split = self.static_split_bash_words(&text);
                    if let Some((first, rest)) = split.split_first() {
                        if let Some(last) = fields.last_mut() {
                            last.push_str(first);
                        }
                        fields.extend(rest.iter().cloned());
                    }
                }
                _ => return Ok(None),
            }
        }
        if !saw_unquoted_split {
            return Ok(None);
        }
        Ok(Some(
            fields
                .into_iter()
                .flat_map(|field| self.static_glob_shell_args(&field))
                .collect(),
        ))
    }

    fn static_mixed_word_shell_args_from_raw(&mut self, raw: &str) -> Option<Vec<ShellArg>> {
        if raw.contains("$(") || raw.contains('`') || !raw.contains('$') {
            return None;
        }
        let chars = raw.chars().collect::<Vec<_>>();
        let mut fields = vec![String::new()];
        let mut saw_unquoted_split = false;
        let mut i = 0usize;
        while i < chars.len() {
            match chars[i] {
                '\'' => {
                    i += 1;
                    while i < chars.len() && chars[i] != '\'' {
                        fields.last_mut()?.push(chars[i]);
                        i += 1;
                    }
                    if i < chars.len() {
                        i += 1;
                    }
                }
                '"' => {
                    i += 1;
                    while i < chars.len() && chars[i] != '"' {
                        if chars[i] == '\\' && i + 1 < chars.len() {
                            let next = chars[i + 1];
                            if matches!(next, '$' | '`' | '"' | '\\' | '\n') {
                                if next != '\n' {
                                    fields.last_mut()?.push(next);
                                }
                                i += 2;
                                continue;
                            }
                        }
                        fields.last_mut()?.push(chars[i]);
                        i += 1;
                    }
                    if i < chars.len() {
                        i += 1;
                    }
                }
                '\\' => {
                    i += 1;
                    if i < chars.len() {
                        fields.last_mut()?.push(chars[i]);
                        i += 1;
                    }
                }
                '$' => {
                    let (name, next_i) = static_dollar_name(&chars, i)?;
                    i = next_i;
                    let value = if name == "@" {
                        self.positional_values
                            .iter()
                            .filter(|value| !value.is_empty())
                            .cloned()
                            .collect::<Vec<_>>()
                            .join(" ")
                    } else if name == "*" {
                        self.positional_values.join(" ")
                    } else {
                        let target = self
                            .resolve_nameref_text(&name)
                            .unwrap_or_else(|| name.clone());
                        self.variable_values.get(&target).cloned()?
                    };
                    saw_unquoted_split = true;
                    let split = self.static_split_bash_words(&value);
                    if let Some((first, rest)) = split.split_first() {
                        fields.last_mut()?.push_str(first);
                        fields.extend(rest.iter().cloned());
                    }
                }
                ch => {
                    fields.last_mut()?.push(ch);
                    i += 1;
                }
            }
        }
        if !saw_unquoted_split {
            return None;
        }
        let mut out = Vec::new();
        for field in fields {
            out.extend(self.static_glob_shell_args(&field));
        }
        Some(out)
    }

    fn static_simple_param_text(&mut self, pair: &Pair<Rule>) -> R<Option<String>> {
        let mut inner = pair.clone().into_inner();
        let Some(inner) = inner.next() else {
            return Ok(Some(String::new()));
        };
        Ok(match inner.as_rule() {
            Rule::name => {
                let target = self
                    .resolve_nameref_text(inner.as_str())
                    .unwrap_or_else(|| inner.as_str().to_string());
                self.variable_values.get(&target).cloned()
            }
            Rule::positional_digit => inner
                .as_str()
                .parse::<usize>()
                .ok()
                .and_then(|n| n.checked_sub(1))
                .and_then(|n| self.positional_values.get(n).cloned()),
            Rule::special_param => literal_string(&special_param_value(inner.as_str()))
                .map(str::to_string),
            _ => None,
        })
    }

    fn static_braced_param_text(&mut self, pair: &Pair<Rule>) -> R<Option<String>> {
        let raw = pair.as_str().trim();
        if let Some(name) = simple_braced_param_name(raw) {
            let target = self
                .resolve_nameref_text(name)
                .unwrap_or_else(|| name.to_string());
            return Ok(self.variable_values.get(&target).cloned());
        }
        Ok(literal_string(&self.braced_param_expr(pair.clone())?).map(str::to_string))
    }

    fn static_split_glob_shell_args(&mut self, value: &str) -> Vec<ShellArg> {
        let mut out = Vec::new();
        for field in self.static_split_bash_words(value) {
            out.extend(self.static_glob_shell_args(&field));
        }
        out
    }

    fn static_glob_shell_args(&mut self, value: &str) -> Vec<ShellArg> {
        if !self.shell_flags.contains(&'f') && word_text_has_glob_meta(value) {
            if let Some(value) = self.path_glob_text_expr(value.to_string()) {
                return vec![ShellArg {
                    value,
                    spread: true,
                }];
            }
        }
        vec![ShellArg {
            value: lit(value),
            spread: false,
        }]
    }

    fn expanded_unquoted_word_shell_args(&mut self, value: Expression) -> Vec<ShellArg> {
        if let Some(text) = literal_string(&value).map(str::to_string) {
            return self.static_split_glob_shell_args(&text);
        }
        vec![ShellArg {
            value: self.split_and_glob_words_expr(value),
            spread: true,
        }]
    }

    fn split_and_glob_words_expr(&mut self, value: Expression) -> Expression {
        let fields = split_bash_words(value);
        if self.shell_flags.contains(&'f') {
            return fields;
        }
        let field = self.fresh("__bash_split_field");
        method(
            fields,
            "flatMap",
            vec![self.dynamic_glob_field_lambda(&field)],
        )
    }

    fn dynamic_glob_field_lambda(&mut self, field: &str) -> Expression {
        let raw = self.fresh("__bash_glob_field");
        let re_src = self.fresh("__bash_glob_re");
        let name = self.fresh("__bash_glob_name");
        let matches = self.fresh("__bash_glob_matches");
        let escaped = method(
            method(
                call_named("__bash_regex_quote", vec![ident(&raw)]),
                "replaceAll",
                vec![lit("\\*"), lit("[\\s\\S]*")],
            ),
            "replaceAll",
            vec![lit("\\?"), lit("[\\s\\S]")],
        );
        let dot_cond = if self.dotglob_enabled {
            Expression::bool(true)
        } else {
            unary(
                UnaryOp::Not,
                method(ident(&name), "startsWith", vec![lit(".")]),
            )
        };
        let filtered = method(
            call_named("__bash_list_dir", vec![bash_path_expr(lit("."))]),
            "filter",
            vec![lambda_expr(
                vec![param_named(&name)],
                binary(
                    BinOp::And,
                    dot_cond,
                    regex_test_expr(regexp(ident(&re_src), ""), ident(&name)),
                ),
            )],
        );
        let no_dynamic_glob = binary(
            BinOp::Or,
            unary(
                UnaryOp::Not,
                regex_test_expr(regexp(lit("[*?]"), ""), ident(&raw)),
            ),
            binary(BinOp::GtEq, method(ident(&raw), "indexOf", vec![lit("/")]), int(0)),
        );
        let no_match_value = if self.nullglob_enabled {
            ident(&matches)
        } else {
            ternary(
                binary(BinOp::StrictEq, member(ident(&matches), "length"), int(0)),
                array(vec![ident(&raw)]),
                ident(&matches),
            )
        };
        lambda_block_with_params(
            vec![param_named(field)],
            vec![
                let_stmt(&raw, call_named("__bash_string", vec![param_value(ident(field))])),
                Statement::new(StmtKind::If {
                    cond: no_dynamic_glob,
                    then_body: vec![Statement::new(StmtKind::Return(Some(array(vec![ident(
                        &raw,
                    )]))))],
                    elifs: Vec::new(),
                    else_body: None,
                }),
                let_stmt(
                    &re_src,
                    parts_to_expr(vec![
                        Part::Text("^".to_string()),
                        Part::Expr(escaped),
                        Part::Text("$".to_string()),
                    ]),
                ),
                let_stmt(&matches, call_named("__bash_sort", vec![filtered])),
                Statement::new(StmtKind::Return(Some(no_match_value))),
            ],
        )
    }

    fn path_glob_word(&mut self, pair: &Pair<Rule>) -> Option<Expression> {
        if self.shell_flags.contains(&'f') {
            return None;
        }
        let text = glob_word_text(pair)?;
        if !word_text_has_glob_meta(&text) {
            return None;
        }
        self.path_glob_text_expr(text)
    }

    fn path_glob_matches_expr(&mut self, text: &str) -> Option<Expression> {
        if self.globstar_enabled {
            if text == "**" {
                return Some(self.globstar_recursive_expr(None, false));
            }
            if text == "**/" {
                return Some(self.globstar_recursive_expr(None, true));
            }
            if let Some(leaf) = text.strip_prefix("**/") {
                if !leaf.contains('/') && word_text_has_glob_meta(leaf) {
                    return Some(self.globstar_recursive_expr(Some(leaf), false));
                }
            }
        }
        if let Some(expr) = self.path_glob_components_expr(text) {
            return Some(expr);
        }
        let (dir, pat, prefix) = match text.rsplit_once('/') {
            Some((dir, pat)) if !dir.is_empty() && !word_text_has_glob_meta(dir) => {
                (dir.to_string(), pat.to_string(), format!("{dir}/"))
            }
            None => (".".to_string(), text.to_string(), String::new()),
            _ => return None,
        };
        let re = glob_to_regex(&pat, false)?;
        let flags = if self.nocaseglob_enabled { "i" } else { "" };
        let item = self.fresh("__bash_glob_name");
        let dot_cond = if self.dotglob_enabled || pat.starts_with('.') {
            Expression::bool(true)
        } else {
            unary(
                UnaryOp::Not,
                method(ident(&item), "startsWith", vec![lit(".")]),
            )
        };
        let list = call_named("__bash_list_dir", vec![bash_path_expr(lit(&dir))]);
        let filtered = method(
            list,
            "filter",
            vec![lambda_expr(
                vec![param_named(&item)],
                binary(
                    BinOp::And,
                    dot_cond,
                    regex_test_expr(regexp(lit(&format!("^(?:{re})$")), flags), ident(&item)),
                ),
            )],
        );
        let matches = if prefix.is_empty() {
            filtered
        } else {
            let map_item = self.fresh("__bash_glob_name");
            array_map_expr(
                filtered,
                &map_item,
                &self.fresh("__bash_glob_out"),
                parts_to_expr(vec![Part::Text(prefix), Part::Expr(ident(&map_item))]),
            )
        };
        Some(call_named("__bash_sort", vec![matches]))
    }

    fn globstar_recursive_expr(&mut self, leaf_pat: Option<&str>, dirs_only: bool) -> Expression {
        let walk = self.fresh("__bash_globstar_walk");
        let path = "__bash_gs_path";
        let cwd = "__bash_gs_cwd";
        let name = "__bash_gs_name";
        let child = "__bash_gs_child";
        let child_expr = ternary(
            binary(BinOp::StrictEq, ident(path), lit(".")),
            ident(name),
            parts_to_expr(vec![
                Part::Expr(ident(path)),
                Part::Text("/".to_string()),
                Part::Expr(ident(name)),
            ]),
        );
        let visible = if self.dotglob_enabled {
            Expression::bool(true)
        } else {
            unary(
                UnaryOp::Not,
                method(ident(name), "startsWith", vec![lit(".")]),
            )
        };
        let dir_cond = call_named(
            "__bash_is_dir",
            vec![bash_path_with_cwd(ident(child), ident(cwd))],
        );
        let descend = call_named(&walk, vec![ident(child), ident(cwd)]);
        let item_result = if dirs_only {
            ternary(
                dir_cond,
                array_spread(vec![
                    (
                        parts_to_expr(vec![
                            Part::Expr(ident(child)),
                            Part::Text("/".to_string()),
                        ]),
                        false,
                    ),
                    (descend, true),
                ]),
                array(Vec::new()),
            )
        } else {
            ternary(
                dir_cond,
                array_spread(vec![(ident(child), false), (descend, true)]),
                array(vec![ident(child)]),
            )
        };
        let walk_decl = let_stmt(
            &walk,
            lambda_block_with_params(
                vec![param_named(path), param_named(cwd)],
                vec![
                    Statement::new(StmtKind::Return(Some(method(
                        call_named(
                            "__bash_list_dir",
                            vec![bash_path_with_cwd(ident(path), ident(cwd))],
                        ),
                        "flatMap",
                        vec![lambda_block_with_params(
                            vec![param_named(name)],
                            vec![
                                let_stmt(child, child_expr),
                                Statement::new(StmtKind::Return(Some(ternary(
                                    visible,
                                    item_result,
                                    array(Vec::new()),
                                )))),
                            ],
                        )],
                    )))),
                ],
            ),
        );
        let mut matches = call_named(&walk, vec![lit("."), ident("PWD")]);
        if let Some(leaf) = leaf_pat {
            let item = self.fresh("__bash_globstar_item");
            let re = glob_to_regex(leaf, false).unwrap_or_else(|| regex_escape(leaf));
            let flags = if self.nocaseglob_enabled { "i" } else { "" };
            matches = method(
                matches,
                "filter",
                vec![lambda_expr(
                    vec![param_named(&item)],
                    binary(
                        BinOp::And,
                        unary(
                            UnaryOp::Not,
                            call_named("__bash_is_dir", vec![bash_path_expr(ident(&item))]),
                        ),
                        regex_test_expr(
                            regexp(lit(&format!("(?:^|/)(?:{re})$")), flags),
                            ident(&item),
                        ),
                    ),
                )],
            );
        }
        iife(vec![
            walk_decl,
            Statement::new(StmtKind::Return(Some(call_named(
                "__bash_sort",
                vec![matches],
            )))),
        ])
    }

    fn path_glob_components_expr(&mut self, text: &str) -> Option<Expression> {
        let (dir_pat, leaf_pat) = text.split_once('/')?;
        if dir_pat.is_empty() || leaf_pat.is_empty() || !word_text_has_glob_meta(dir_pat) {
            return None;
        }
        if leaf_pat.contains('/') {
            return None;
        }
        let dirs = self.path_glob_matches_expr(dir_pat)?;
        let leaf_re = glob_to_regex(leaf_pat, false)?;
        let flags = if self.nocaseglob_enabled { "i" } else { "" };
        let dir = self.fresh("__bash_glob_dir");
        let leaf = self.fresh("__bash_glob_leaf");
        let dot_cond = if self.dotglob_enabled || leaf_pat.starts_with('.') {
            Expression::bool(true)
        } else {
            unary(
                UnaryOp::Not,
                method(ident(&leaf), "startsWith", vec![lit(".")]),
            )
        };
        let leaf_matches = method(
            call_named("__bash_list_dir", vec![bash_path_expr(ident(&dir))]),
            "filter",
            vec![lambda_expr(
                vec![param_named(&leaf)],
                binary(
                    BinOp::And,
                    dot_cond,
                    regex_test_expr(
                        regexp(lit(&format!("^(?:{leaf_re})$")), flags),
                        ident(&leaf),
                    ),
                ),
            )],
        );
        let prefixed = array_map_expr(
            leaf_matches,
            &leaf,
            &self.fresh("__bash_glob_leaf_out"),
            parts_to_expr(vec![
                Part::Expr(ident(&dir)),
                Part::Text("/".to_string()),
                Part::Expr(ident(&leaf)),
            ]),
        );
        let per_dir = ternary(
            call_named("__bash_is_dir", vec![bash_path_expr(ident(&dir))]),
            prefixed,
            array(Vec::new()),
        );
        Some(call_named(
            "__bash_sort",
            vec![method(
                dirs,
                "flatMap",
                vec![lambda_expr(vec![param_named(&dir)], per_dir)],
            )],
        ))
    }

    fn path_glob_text_expr(&mut self, text: String) -> Option<Expression> {
        let matches = self.path_glob_matches_expr(&text)?;
        if self.nullglob_enabled || self.failglob_enabled {
            return Some(matches);
        }
        let tmp = self.fresh("__bash_glob_matches");
        Some(iife(vec![
            let_stmt(&tmp, matches),
            Statement::new(StmtKind::Return(Some(ternary(
                binary(BinOp::StrictEq, member(ident(&tmp), "length"), int(0)),
                array(vec![lit(&text)]),
                ident(&tmp),
            )))),
        ]))
    }

    fn apply_failglob_guards(
        &mut self,
        body: Vec<Statement>,
        words: &[Pair<Rule>],
    ) -> Vec<Statement> {
        if !self.failglob_enabled || self.shell_flags.contains(&'f') {
            return body;
        }
        let mut checks = Vec::new();
        for word in words {
            if let Some(text) = self.failglob_pattern_text(word) {
                if word_text_has_glob_meta(&text) {
                    if let Some(matches) = self.path_glob_matches_expr(&text) {
                        checks.push((text, matches));
                    }
                }
            }
        }
        let mut wrapped = body;
        for (text, matches) in checks.into_iter().rev() {
            wrapped = vec![Statement::new(StmtKind::If {
                cond: binary(BinOp::StrictEq, member(matches, "length"), int(0)),
                then_body: vec![
                    bash_stderr_stmt(call_named(
                        "__bash_sprintf",
                        vec![lit("bash: no match: %s\n"), lit(&text)],
                    )),
                    assign_stmt(ident("__bash_abort_line"), Expression::bool(true)),
                    assign_stmt(ident("__bash_status"), int(1)),
                ],
                elifs: Vec::new(),
                else_body: Some(wrapped),
            })];
        }
        wrapped
    }

    fn failglob_pattern_text(&self, pair: &Pair<Rule>) -> Option<String> {
        if let Some(text) = glob_word_text(pair) {
            return Some(text);
        }
        let raw = pair.as_str().trim();
        if raw.starts_with('"') || raw.starts_with('\'') {
            return None;
        }
        let name = whole_unquoted_param_name(raw)?;
        let target = self
            .resolve_nameref_text(name)
            .unwrap_or_else(|| name.to_string());
        self.variable_values.get(&target).cloned()
    }

    fn assignment_word_as_text(&mut self, pair: Pair<Rule>) -> R<Expression> {
        let mut parts = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::assign_target | Rule::assign_op => {
                    parts.push(Part::Text(p.as_str().to_string()))
                }
                Rule::word | Rule::assignment_value_word => parts.extend(self.word_parts(p)?),
                Rule::array_literal => parts.push(Part::Text(p.as_str().to_string())),
                _ => {}
            }
        }
        Ok(parts_to_expr(parts))
    }

    /// A word as exactly one value (brace expansion joined back if it produced several).
    fn word_expr(&mut self, pair: Pair<Rule>) -> R<Expression> {
        let mut vals = self.expand_word(pair)?;
        if vals.len() == 1 {
            return Ok(vals.pop().unwrap());
        }
        if vals.is_empty() {
            return Ok(lit(""));
        }
        Ok(array_join(array(vals), lit(" ")))
    }

    /// A word as the list of words it expands to — brace expansion may yield many.
    fn expand_word(&mut self, pair: Pair<Rule>) -> R<Vec<Expression>> {
        if !word_has_quoted_part(&pair)
            && self
                .static_word_text(&pair)
                .is_some_and(|text| text.is_empty())
        {
            return Ok(Vec::new());
        }
        if self.brace_expansion_enabled && !word_has_quoted_part(&pair) {
            if let Some(expanded_words) = brace_expand_text(pair.as_str()) {
                let saved = self.brace_expansion_enabled;
                self.brace_expansion_enabled = false;
                let mut out = Vec::new();
                for text in expanded_words {
                    if text.is_empty() {
                        out.push(lit(""));
                        continue;
                    }
                    let mut parsed = WashmParser::parse(Rule::word, &text)
                        .map_err(|e| format!("brace expansion: {e}"))?;
                    let Some(word) = parsed.next() else {
                        continue;
                    };
                    out.extend(self.expand_word(word)?);
                }
                self.brace_expansion_enabled = saved;
                return Ok(out);
            }
        }
        if let Some(keys) = self.exact_indirect_array_key_values(&pair) {
            return Ok(keys);
        }
        if let Some(names) = self.whole_word_prefix_names(pair.as_str()) {
            return Ok(names.into_iter().map(|name| lit(&name)).collect());
        }
        // Each part is a list of alternatives; the word is their cartesian product.
        let mut alternatives: Vec<Vec<Vec<Part>>> = Vec::new();
        for p in pair.into_inner() {
            if self.brace_expansion_enabled && p.as_rule() == Rule::brace_expansion {
                alternatives.push(self.brace_alternatives(p)?);
            } else {
                alternatives.push(vec![self.part_to_parts(p)?]);
            }
        }
        let mut combos: Vec<Vec<Part>> = vec![Vec::new()];
        for alts in alternatives {
            let mut next = Vec::new();
            for c in &combos {
                for a in &alts {
                    let mut n = c.clone();
                    n.extend(a.iter().cloned());
                    next.push(n);
                }
            }
            combos = next;
        }
        Ok(combos.into_iter().map(parts_to_expr).collect())
    }

    fn whole_word_prefix_names(&self, text: &str) -> Option<Vec<String>> {
        let body = if text.starts_with('"') && text.ends_with('"') && text.len() >= 2 {
            &text[1..text.len() - 1]
        } else {
            text
        };
        let inner = body.strip_prefix("${!")?.strip_suffix('}')?;
        let prefix = inner
            .strip_suffix('@')
            .or_else(|| inner.strip_suffix('*'))?;
        if !prefix
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '_')
        {
            return None;
        }
        Some(
            self.variables
                .iter()
                .filter(|name| name.starts_with(prefix))
                .cloned()
                .collect(),
        )
    }

    fn word_parts(&mut self, pair: Pair<Rule>) -> R<Vec<Part>> {
        let mut parts = Vec::new();
        for p in pair.into_inner() {
            parts.extend(self.part_to_parts(p)?);
        }
        Ok(parts)
    }

    /// `{a,b}` / `{1..3}` → the alternatives, each a part list.
    fn brace_alternatives(&mut self, pair: Pair<Rule>) -> R<Vec<Vec<Part>>> {
        let inner = pair.into_inner().next().ok_or("empty brace expansion")?;
        match inner.as_rule() {
            Rule::brace_range => {
                let ends: Vec<&str> = inner.clone().into_inner().map(|p| p.as_str()).collect();
                if ends.len() < 2 {
                    return Err("malformed brace range".into());
                }
                let (a, b) = (ends[0], ends[1]);
                let step: i64 = ends.get(2).and_then(|s| s.parse().ok()).unwrap_or(1);
                let mut out = Vec::new();
                if let (Ok(x), Ok(y)) = (a.parse::<i64>(), b.parse::<i64>()) {
                    let width = if (a.starts_with('0') && a.len() > 1)
                        || (b.starts_with('0') && b.len() > 1)
                    {
                        a.len().max(b.len())
                    } else {
                        0
                    };
                    let step = step.abs().max(1);
                    let mut v = x;
                    loop {
                        if (x <= y && v > y) || (x > y && v < y) {
                            break;
                        }
                        out.push(vec![Part::Text(if width > 0 {
                            format!("{v:0width$}")
                        } else {
                            v.to_string()
                        })]);
                        v += if x <= y { step } else { -step };
                    }
                } else {
                    let (x, y) = (
                        a.chars().next().unwrap() as u32,
                        b.chars().next().unwrap() as u32,
                    );
                    let step = step.unsigned_abs().max(1) as u32;
                    let mut v = x;
                    loop {
                        if (x <= y && v > y) || (x > y && v < y) {
                            break;
                        }
                        out.push(vec![Part::Text(
                            char::from_u32(v).unwrap_or('?').to_string(),
                        )]);
                        if x <= y {
                            v += step
                        } else {
                            v = v.wrapping_sub(step)
                        };
                    }
                }
                Ok(out)
            }
            Rule::brace_sequence => {
                // Elements separated by commas; an absent element is the empty string.
                let text = inner.as_str();
                let mut out = Vec::new();
                let mut elements: Vec<Option<Pair<Rule>>> = Vec::new();
                let mut expect_element = true;
                let mut pos = 0usize;
                let inner_start = inner.as_span().start();
                for p in inner.clone().into_inner() {
                    let start = p.as_span().start() - inner_start;
                    for ch in text[pos..start].chars() {
                        if ch == ',' {
                            if expect_element {
                                elements.push(None);
                            }
                            expect_element = true;
                        }
                    }
                    elements.push(Some(p.clone()));
                    expect_element = false;
                    pos = p.as_span().end() - inner_start;
                }
                for ch in text[pos..].chars() {
                    if ch == ',' {
                        if expect_element {
                            elements.push(None);
                        }
                        expect_element = true;
                    }
                }
                if expect_element {
                    elements.push(None);
                }
                for el in elements {
                    match el {
                        None => out.push(vec![Part::Text(String::new())]),
                        Some(p) => {
                            // A nested brace expansion multiplies the element.
                            let mut alts: Vec<Vec<Part>> = vec![Vec::new()];
                            for part in p.into_inner() {
                                let part_alts = if part.as_rule() == Rule::brace_expansion {
                                    self.brace_alternatives(part)?
                                } else {
                                    vec![self.part_to_parts(part)?]
                                };
                                let mut next = Vec::new();
                                for c in &alts {
                                    for a in &part_alts {
                                        let mut n = c.clone();
                                        n.extend(a.iter().cloned());
                                        next.push(n);
                                    }
                                }
                                alts = next;
                            }
                            out.extend(alts);
                        }
                    }
                }
                Ok(out)
            }
            _ => Err("malformed brace expansion".into()),
        }
    }

    fn part_to_parts(&mut self, pair: Pair<Rule>) -> R<Vec<Part>> {
        Ok(match pair.as_rule() {
            Rule::bare_word
            | Rule::assignment_bare_word
            | Rule::brace_seq_text
            | Rule::pattern_text
            | Rule::brace_text => vec![Part::Text(unescape_bare(pair.as_str()))],
            Rule::brace_expansion => vec![Part::Text(pair.as_str().to_string())],
            Rule::close_brace_arg | Rule::close_brace_tail => vec![Part::Text("}".to_string())],
            Rule::extglob_word => vec![Part::Text(pair.as_str().to_string())],
            Rule::single_quoted_string => {
                let s = pair.as_str();
                vec![Part::Text(s[1..s.len() - 1].to_string())]
            }
            Rule::ansi_c_quoting => {
                let s = pair.as_str();
                vec![Part::Text(decode_ansi_c(&s[2..s.len() - 1]))]
            }
            Rule::quoted_string => self.quoted_parts(pair)?,
            Rule::locale_quoting => {
                let mut v = Vec::new();
                for p in pair.into_inner() {
                    v.extend(self.quoted_parts(p)?);
                }
                v
            }
            Rule::tilde_expansion => vec![Part::Expr(self.tilde_expr(pair)?)],
            _ => vec![Part::Expr(self.part_expr(pair)?)],
        })
    }

    fn quoted_parts(&mut self, pair: Pair<Rule>) -> R<Vec<Part>> {
        let mut parts = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::dq_text => parts.push(Part::Text(p.as_str().to_string())),
                Rule::dq_escape => {
                    let c = p.as_str().chars().nth(1).unwrap_or('\\');
                    parts.push(Part::Text(match c {
                        '$' | '`' | '"' | '\\' => c.to_string(),
                        '\n' => String::new(),
                        _ => format!("\\{c}"),
                    }));
                }
                Rule::braced_param => {
                    parts.push(Part::Expr(self.braced_param_expr_with_tilde(p, false)?))
                }
                _ => parts.push(Part::Expr(self.part_expr(p)?)),
            }
        }
        Ok(parts)
    }

    /// An expansion part as one expression.
    fn part_expr(&mut self, pair: Pair<Rule>) -> R<Expression> {
        match pair.as_rule() {
            Rule::simple_param => {
                let inner = pair.into_inner().next().ok_or("empty parameter")?;
                Ok(match inner.as_rule() {
                    Rule::name if inner.as_str() == "BASH_SUBSHELL" => {
                        lit(&self.subshell_depth.to_string())
                    }
                    Rule::name => param_value(self.name_or_element_target(inner.as_str())?),
                    Rule::positional_digit => param_value(positional(inner.as_str())),
                    _ => special_param_value(inner.as_str()),
                })
            }
            Rule::braced_param => self.braced_param_expr(pair),
            Rule::dollar_paren_subst => {
                let mut list = None;
                for p in pair.into_inner() {
                    if p.as_rule() == Rule::list {
                        list = Some(p);
                    }
                }
                let Some(list) = list else { return Ok(lit("")) };
                // `$(< file)` reads the file without running a command.
                if let Some(target) = self.read_only_redirection(&list) {
                    let t = self.word_expr(target)?;
                    return Ok(strip_command_substitution_value(
                        self.read_file_or_fd_expr(t),
                    ));
                }
                if let Some(e) = self.simple_command_substitution_value(&list)? {
                    return Ok(command_substitution_expr(e));
                }
                let mut child = self.clone();
                child.subshell_depth += 1;
                let body = child.walk_list(list)?;
                self.heredocs = child.heredocs.clone();
                self.counter = child.counter;
                let (params, args) = self.shell_child_bindings(&child, true, &[], None);
                Ok(command_substitution_expr(capture_with_args(
                    body, params, args,
                )))
            }
            Rule::brace_command_subst => {
                let reply_form = pair.as_str().starts_with("${|");
                let mut list = None;
                for p in pair.into_inner() {
                    if p.as_rule() == Rule::list {
                        list = Some(p);
                    }
                }
                let body = match list {
                    Some(list) => self.walk_list(list)?,
                    None => Vec::new(),
                };
                if reply_form {
                    Ok(sequence(vec![iife(body), param_value(ident("REPLY"))]))
                } else {
                    Ok(command_substitution_expr(current_shell_capture(
                        body,
                        self.subshell_depth > 0,
                    )))
                }
            }
            Rule::backtick_subst => {
                let s = pair.as_str();
                let inner = unescape_backtick(&s[1..s.len() - 1]);
                let (stripped, heredocs) = extract_heredocs(&inner);
                let pairs = WashmParser::parse(Rule::program, &stripped)
                    .map_err(|e| format!("backtick substitution: {e}"))?;
                let mut child = self.clone();
                child.heredocs = heredocs.into();
                child.subshell_depth += 1;
                let mut body = Vec::new();
                for p in pairs {
                    if p.as_rule() == Rule::program {
                        for inner in p.into_inner() {
                            if inner.as_rule() == Rule::list {
                                body.extend(child.walk_list(inner)?);
                            }
                        }
                    }
                }
                self.heredocs = child.heredocs.clone();
                self.counter = child.counter;
                let (params, args) = self.shell_child_bindings(&child, true, &[], None);
                Ok(command_substitution_expr(capture_with_args(
                    body, params, args,
                )))
            }
            Rule::arithmetic_expansion => Ok(arith_string(self.arith_expansion_expr(pair)?)),
            Rule::process_substitution => {
                let mut dir = "<".to_string();
                let mut list = None;
                for p in pair.into_inner() {
                    match p.as_rule() {
                        Rule::procsub_dir => dir = p.as_str().to_string(),
                        Rule::procsub_list => {
                            for l in p.into_inner() {
                                if l.as_rule() == Rule::list {
                                    list = Some(l);
                                }
                            }
                        }
                        _ => {}
                    }
                }
                if dir == "<" {
                    let mut child = self.clone();
                    child.subshell_depth += 1;
                    let body = match list {
                        Some(list) => child.walk_list(list)?,
                        None => Vec::new(),
                    };
                    self.heredocs = child.heredocs.clone();
                    self.counter = child.counter;
                    let (params, args) = self.shell_child_bindings(&child, true, &[], None);
                    let fd_key = (64 + self.counter).to_string();
                    self.counter += 1;
                    let path = format!("/dev/fd/{fd_key}");
                    return Ok(sequence(vec![
                        assign_expr(
                            ident("__bash_job_seq"),
                            binary(BinOp::Add, ident("__bash_job_seq"), int(1)),
                        ),
                        assign_expr(ident("__bash_last_bg_pid"), ident("__bash_job_seq")),
                        assign_expr(
                            index(ident("__bash_fds"), lit(&fd_key)),
                            capture_with_args(body, params, args),
                        ),
                        assign_expr(
                            index(ident("__bash_jobs"), ident("__bash_last_bg_pid")),
                            ident("__bash_status"),
                        ),
                        lit(&path),
                    ]));
                }
                let mut child = self.clone();
                child.subshell_depth += 1;
                let body = match list {
                    Some(list) => child.walk_list(list)?,
                    None => Vec::new(),
                };
                self.heredocs = child.heredocs.clone();
                self.counter = child.counter;
                let fd_key = (64 + self.counter).to_string();
                self.counter += 1;
                let path = format!("/dev/fd/{fd_key}");
                let mut procsub_body = child
                    .variables
                    .difference(&self.variables)
                    .filter(|name| is_name(name) && name.as_str() != BASH_ARGS)
                    .map(|name| let_stmt(name, undefined()))
                    .collect::<Vec<_>>();
                procsub_body.extend(body);
                procsub_body.push(Statement::new(StmtKind::Return(Some(ident(
                    "__bash_status",
                )))));
                Ok(sequence(vec![
                    assign_expr(
                        ident("__bash_job_seq"),
                        binary(BinOp::Add, ident("__bash_job_seq"), int(1)),
                    ),
                    assign_expr(ident("__bash_last_bg_pid"), ident("__bash_job_seq")),
                    assign_expr(
                        index(ident("__bash_coprocs"), lit(&fd_key)),
                        object(vec![
                            ("read_fd", lit(&fd_key)),
                            ("pid", ident("__bash_last_bg_pid")),
                            ("stdout", lit("parent")),
                            (
                                "body",
                                lambda_block_with_params(
                                    vec![param_named("__bash_stdin")],
                                    procsub_body,
                                ),
                            ),
                        ]),
                    ),
                    assign_expr(
                        index(ident("__bash_jobs"), ident("__bash_last_bg_pid")),
                        int(0),
                    ),
                    lit(&path),
                ]))
            }
            Rule::tilde_expansion => self.tilde_expr(pair),
            Rule::quoted_string
            | Rule::single_quoted_string
            | Rule::ansi_c_quoting
            | Rule::locale_quoting
            | Rule::bare_word
            | Rule::assignment_bare_word
            | Rule::close_brace_arg
            | Rule::close_brace_tail => Ok(parts_to_expr(self.part_to_parts(pair)?)),
            Rule::word | Rule::assignment_value_word => self.word_expr(pair),
            other => Err(format!(
                "unsupported word part {other:?}: {}",
                pair.as_str()
            )),
        }
    }

    /// `$(< file)`: the list is one simple command made only of a `<` redirection.
    fn read_only_redirection<'i>(&mut self, list: &Pair<'i, Rule>) -> Option<Pair<'i, Rule>> {
        let and_ors: Vec<Pair<Rule>> = list
            .clone()
            .into_inner()
            .filter(|p| p.as_rule() == Rule::and_or)
            .collect();
        if and_ors.len() != 1 {
            return None;
        }
        let pipelines: Vec<Pair<Rule>> = and_ors[0]
            .clone()
            .into_inner()
            .filter(|p| p.as_rule() == Rule::pipeline)
            .collect();
        if pipelines.len() != 1 {
            return None;
        }
        let cmds: Vec<Pair<Rule>> = pipelines[0].clone().into_inner().collect();
        if cmds.len() != 1 || cmds[0].as_rule() != Rule::simple_command {
            return None;
        }
        let children: Vec<Pair<Rule>> = cmds[0].clone().into_inner().collect();
        if children.len() != 1 || children[0].as_rule() != Rule::redirection {
            return None;
        }
        let fd = children[0].clone().into_inner().next()?;
        if fd.as_rule() != Rule::fd_redirection {
            return None;
        }
        let mut op = String::new();
        let mut target = None;
        for p in fd.into_inner() {
            match p.as_rule() {
                Rule::redir_op => op = p.as_str().to_string(),
                Rule::redir_target => target = p.into_inner().next(),
                _ => {}
            }
        }
        if op != "<" {
            return None;
        }
        target.filter(|t| t.as_rule() == Rule::word)
    }

    fn simple_command_substitution_value(&mut self, list: &Pair<Rule>) -> R<Option<Expression>> {
        let and_ors: Vec<Pair<Rule>> = list
            .clone()
            .into_inner()
            .filter(|p| p.as_rule() == Rule::and_or)
            .collect();
        if and_ors.len() != 1 {
            return Ok(None);
        }
        let pipelines: Vec<Pair<Rule>> = and_ors[0]
            .clone()
            .into_inner()
            .filter(|p| p.as_rule() == Rule::pipeline)
            .collect();
        if pipelines.len() != 1 {
            return Ok(None);
        }
        let (_, _, units) = self.pipeline_units(pipelines[0].clone());
        if units.len() != 1
            || !units[0].redirs.is_empty()
            || units[0].cmd.as_rule() != Rule::simple_command
        {
            return Ok(None);
        }
        let Some((name, suffix)) = simple_command_literal_parts(units[0].cmd.clone()) else {
            return Ok(None);
        };
        if self.expand_aliases_enabled && self.aliases.contains_key(&name) {
            return Ok(None);
        }
        match name.as_str() {
            _ if is_function_name(&name) && self.functions.contains_key(&name) => {
                let lowered = if let Some(inlined) = self.inline_function_call(&name, &suffix)? {
                    inlined
                } else {
                    let mut child = self.clone();
                    child.subshell_depth += 1;
                    let Some(lowered) = child.user_function_call(&name, &suffix)? else {
                        return Ok(None);
                    };
                    self.counter = child.counter;
                    lowered
                };
                Ok(Some(capture(lowered.into_stmts())))
            }
            "printf" => self.printf_value(suffix),
            "echo" => {
                let args = self.words_shell_args(suffix)?;
                Ok(Some(array_join(shell_args_array(args), lit(" "))))
            }
            _ if !BASH_BUILTIN_NAMES.contains(&name.as_str()) => {
                let args = self.words_shell_args(suffix)?;
                let lowered = self.external_command(&name, args);
                Ok(Some(capture(lowered.into_stmts())))
            }
            _ => Ok(None),
        }
    }

    fn tilde_expr(&mut self, pair: Pair<Rule>) -> R<Expression> {
        let mut base = param_value(ident("HOME"));
        let mut special: Option<String> = None;
        let mut path = String::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::tilde_special => {
                    special = Some(p.as_str().to_string());
                    base = if p.as_str() == "+" {
                        param_value(ident("PWD"))
                    } else {
                        ternary(
                            binary(BinOp::StrictEq, ident("OLDPWD"), undefined()),
                            lit("~-"),
                            ident("OLDPWD"),
                        )
                    };
                }
                Rule::tilde_user => {
                    let user = p.as_str();
                    base = ternary(
                        binary(
                            BinOp::StrictEq,
                            bash_env_value_expr("USER", lit("")),
                            lit(user),
                        ),
                        param_value(ident("HOME")),
                        lit(&format!("~{user}")),
                    );
                }
                Rule::tilde_path => path = unescape_bare(p.as_str()),
                _ => {}
            }
        }
        if path.is_empty() {
            return Ok(base);
        }
        if special.as_deref() == Some("-") {
            base = ternary(
                binary(BinOp::StrictEq, ident("OLDPWD"), undefined()),
                lit("~-"),
                ident("OLDPWD"),
            );
        }
        Ok(parts_to_expr(vec![Part::Expr(base), Part::Text(path)]))
    }

    fn exact_transformed_array_word(&mut self, pair: &Pair<Rule>) -> R<Option<Expression>> {
        let Some(braced) = exact_braced_param_word(pair) else {
            return Ok(None);
        };
        let inner = braced.into_inner().next().ok_or("empty ${}")?;
        match inner.as_rule() {
            Rule::search_replace => {
                let (expr, all) = self.search_replace_expansion(inner, false)?;
                Ok(all.then_some(expr))
            }
            Rule::prefix_suffix_removal => {
                let (expr, all) = self.prefix_suffix_expansion(inner, false)?;
                Ok(all.then_some(expr))
            }
            Rule::case_modification => {
                let (expr, all) = self.case_mod_expansion(inner, false)?;
                Ok(all.then_some(expr))
            }
            Rule::transform_expansion => {
                let (expr, all) = self.transform_expansion(inner, false)?;
                Ok(all.then_some(expr))
            }
            _ => Ok(None),
        }
    }

    fn search_replace_expansion(
        &mut self,
        inner: Pair<Rule>,
        join_all: bool,
    ) -> R<(Expression, bool)> {
        let mut target = None;
        let mut all = false;
        let mut op = String::new();
        let mut pattern: Option<Pair<Rule>> = None;
        let mut replacement: Option<Pair<Rule>> = None;
        for p in inner.into_inner() {
            match p.as_rule() {
                Rule::param_base => {
                    let (t, a) = self.param_base_parts(p)?;
                    target = Some(t);
                    all = a;
                }
                Rule::search_replace_op => op = p.as_str().to_string(),
                Rule::glob_pattern => pattern = Some(p),
                Rule::replacement_word => replacement = Some(p),
                _ => {}
            }
        }
        let target = target.ok_or("replace without parameter")?;
        let literal_pat = pattern
            .as_ref()
            .and_then(|p| self.static_glob_pattern_text(p));
        let static_rep = replacement
            .as_ref()
            .and_then(|p| self.static_replacement_text(p))
            .map(|s| bash_replacement_text(&s, self.patsub_replacement_enabled));
        let rep_expr = match replacement.clone() {
            Some(p) => match static_rep {
                Some(s) => lit(&s),
                None => self.brace_word_expr(p)?,
            },
            None => lit(""),
        };
        let make_one = |this: &mut Self, value: Expression| {
            this.search_replace_value_expr(
                param_value(value),
                &op,
                literal_pat.as_deref(),
                pattern.clone(),
                rep_expr.clone(),
            )
        };
        let expr = if all {
            let item = self.fresh("__bash_param_item");
            let out = self.fresh("__bash_param_out");
            let mapped = make_one(self, ident(&item))?;
            let array = array_map_expr(target, &item, &out, mapped);
            if join_all {
                array_join(array, ifs_join_sep())
            } else {
                array
            }
        } else {
            make_one(self, target)?
        };
        Ok((expr, all))
    }

    fn search_replace_value_expr(
        &mut self,
        target: Expression,
        op: &str,
        literal_pat: Option<&str>,
        pattern: Option<Pair<Rule>>,
        rep_expr: Expression,
    ) -> R<Expression> {
        if let Some(pat) = literal_pat {
            if let Some(re) = glob_to_regex(pat, false) {
                let (src, mut flags) = match op {
                    "//" if pat == "*" => (format!("^(?:{re})$"), String::new()),
                    "//" => (re, String::from("g")),
                    "/#" => (format!("^(?:{re})"), String::new()),
                    "/%" => (format!("(?:{re})$"), String::new()),
                    _ => (re, String::new()),
                };
                if self.nocasematch_enabled {
                    flags.push('i');
                }
                return Ok(regex_replace_expr(target, &src, &flags, rep_expr));
            }
        }
        let pat_expr = match pattern {
            Some(p) => self.brace_word_expr(p)?,
            None => lit(""),
        };
        Ok(match op {
            "//" => method(target, "replaceAll", vec![pat_expr, rep_expr]),
            "/#" => ternary(
                method(target.clone(), "startsWith", vec![pat_expr.clone()]),
                parts_to_expr(vec![
                    Part::Expr(rep_expr),
                    Part::Expr(method(
                        target.clone(),
                        "slice",
                        vec![member(pat_expr, "length")],
                    )),
                ]),
                target,
            ),
            "/%" => ternary(
                method(target.clone(), "endsWith", vec![pat_expr.clone()]),
                parts_to_expr(vec![
                    Part::Expr(method(
                        target.clone(),
                        "slice",
                        vec![
                            int(0),
                            binary(
                                BinOp::Sub,
                                member(target.clone(), "length"),
                                member(pat_expr, "length"),
                            ),
                        ],
                    )),
                    Part::Expr(rep_expr),
                ]),
                target,
            ),
            _ => method(target, "replace", vec![pat_expr, rep_expr]),
        })
    }

    fn prefix_suffix_expansion(
        &mut self,
        inner: Pair<Rule>,
        join_all: bool,
    ) -> R<(Expression, bool)> {
        let mut target = None;
        let mut all = false;
        let mut op = String::new();
        let mut pattern: Option<Pair<Rule>> = None;
        for p in inner.into_inner() {
            match p.as_rule() {
                Rule::param_base => {
                    let (t, a) = self.param_base_parts(p)?;
                    target = Some(t);
                    all = a;
                }
                Rule::removal_op => op = p.as_str().to_string(),
                Rule::removal_pattern => pattern = Some(p),
                _ => {}
            }
        }
        let target = target.ok_or("removal without parameter")?;
        let Some(pattern) = pattern else {
            return Ok((
                if all && join_all {
                    array_join(target, ifs_join_sep())
                } else {
                    param_value(target)
                },
                all,
            ));
        };
        let literal_pat = self.static_glob_pattern_text(&pattern);
        let make_one = |this: &mut Self, value: Expression| {
            this.prefix_suffix_value_expr(
                param_value(value),
                &op,
                literal_pat.as_deref(),
                pattern.clone(),
            )
        };
        let expr = if all {
            let item = self.fresh("__bash_param_item");
            let out = self.fresh("__bash_param_out");
            let mapped = make_one(self, ident(&item))?;
            let array = array_map_expr(target, &item, &out, mapped);
            if join_all {
                array_join(array, ifs_join_sep())
            } else {
                array
            }
        } else {
            make_one(self, target)?
        };
        Ok((expr, all))
    }

    fn prefix_suffix_value_expr(
        &mut self,
        target: Expression,
        op: &str,
        literal_pat: Option<&str>,
        pattern: Pair<Rule>,
    ) -> R<Expression> {
        if let Some(pat) = literal_pat {
            let shortest = op == "#" || op == "%";
            if let Some(re) = glob_to_regex(pat, shortest && op == "#") {
                return Ok(match op {
                    "#" | "##" => regex_replace_expr(target, &format!("^(?:{re})"), "", lit("")),
                    "%" => {
                        regex_replace_expr(target, &format!("^([\\s\\S]*)(?:{re})$"), "", lit("$1"))
                    }
                    _ => regex_replace_expr(
                        target,
                        &format!("^([\\s\\S]*?)(?:{re})$"),
                        "",
                        lit("$1"),
                    ),
                });
            }
        }
        let pat_expr = self.brace_word_expr(pattern)?;
        Ok(match op {
            "#" | "##" => ternary(
                method(target.clone(), "startsWith", vec![pat_expr.clone()]),
                method(target.clone(), "slice", vec![member(pat_expr, "length")]),
                target,
            ),
            _ => ternary(
                method(target.clone(), "endsWith", vec![pat_expr.clone()]),
                method(
                    target.clone(),
                    "slice",
                    vec![
                        int(0),
                        binary(
                            BinOp::Sub,
                            member(target.clone(), "length"),
                            member(pat_expr, "length"),
                        ),
                    ],
                ),
                target,
            ),
        })
    }

    fn case_mod_expansion(&mut self, inner: Pair<Rule>, join_all: bool) -> R<(Expression, bool)> {
        let mut target = None;
        let mut all = false;
        let mut op = String::new();
        let mut pattern: Option<Pair<Rule>> = None;
        for p in inner.into_inner() {
            match p.as_rule() {
                Rule::param_base => {
                    let (t, a) = self.param_base_parts(p)?;
                    target = Some(t);
                    all = a;
                }
                Rule::case_mod_op => op = p.as_str().to_string(),
                Rule::removal_pattern => pattern = Some(p),
                _ => {}
            }
        }
        let target = target.ok_or("case modification without parameter")?;
        let literal_pat = pattern
            .as_ref()
            .and_then(|p| self.static_glob_pattern_text(p));
        let make_one = |this: &mut Self, value: Expression| {
            this.case_mod_value_expr(
                param_value(value),
                &op,
                literal_pat.as_deref(),
                pattern.clone(),
            )
        };
        let expr = if all {
            let item = self.fresh("__bash_param_item");
            let out = self.fresh("__bash_param_out");
            let mapped = make_one(self, ident(&item))?;
            let array = array_map_expr(target, &item, &out, mapped);
            if join_all {
                array_join(array, ifs_join_sep())
            } else {
                array
            }
        } else {
            make_one(self, target)?
        };
        Ok((expr, all))
    }

    fn case_mod_value_expr(
        &mut self,
        target: Expression,
        op: &str,
        literal_pat: Option<&str>,
        pattern: Option<Pair<Rule>>,
    ) -> R<Expression> {
        let case_method = if op.starts_with(',') {
            "toLowerCase"
        } else {
            "toUpperCase"
        };
        if pattern.is_none() {
            return Ok(match op {
                "^^" => method(target, "toUpperCase", vec![]),
                ",," => method(target, "toLowerCase", vec![]),
                "~~" => regex_replace_expr(target, "[\\s\\S]", "g", case_toggle_lambda()),
                "~" => regex_replace_expr(target, "[\\s\\S]", "", case_toggle_lambda()),
                "^" => parts_to_expr(vec![
                    Part::Expr(method(
                        method(target.clone(), "charAt", vec![int(0)]),
                        "toUpperCase",
                        vec![],
                    )),
                    Part::Expr(method(target, "slice", vec![int(1)])),
                ]),
                "," => parts_to_expr(vec![
                    Part::Expr(method(
                        method(target.clone(), "charAt", vec![int(0)]),
                        "toLowerCase",
                        vec![],
                    )),
                    Part::Expr(method(target, "slice", vec![int(1)])),
                ]),
                _ => target,
            });
        }
        let pat = match literal_pat {
            Some("") | None => "[\\s\\S]".to_string(),
            Some("?") | Some("*") => "[\\s\\S]".to_string(),
            Some(text) => glob_to_regex(text, false).unwrap_or_else(|| regex_escape(text)),
        };
        let replacement = if op.starts_with('~') {
            case_toggle_lambda()
        } else {
            lambda_expr(
                vec![param_named("__bash_ch")],
                method(ident("__bash_ch"), case_method, vec![]),
            )
        };
        if !matches!(op, "^^" | ",," | "~~") {
            let first = method(target.clone(), "charAt", vec![int(0)]);
            let rest = method(target.clone(), "slice", vec![int(1)]);
            let matches_first =
                regex_test_expr(regexp(lit(&format!("^(?:{pat})$")), ""), first.clone());
            let converted = if op.starts_with('~') {
                call(case_toggle_lambda(), vec![first])
            } else {
                method(first, case_method, vec![])
            };
            return Ok(ternary(
                matches_first,
                parts_to_expr(vec![Part::Expr(converted), Part::Expr(rest)]),
                target,
            ));
        }
        let flags = "g";
        Ok(regex_replace_expr(target, &pat, flags, replacement))
    }

    fn transform_expansion(&mut self, inner: Pair<Rule>, join_all: bool) -> R<(Expression, bool)> {
        let mut target = None;
        let mut target_name = None;
        let mut all = false;
        let mut op = String::new();
        for p in inner.into_inner() {
            match p.as_rule() {
                Rule::param_base => {
                    target_name = param_base_name(&p);
                    let (t, a) = self.param_base_parts(p)?;
                    target = Some(t);
                    all = a;
                }
                Rule::transform_op => op = p.as_str().to_string(),
                _ => {}
            }
        }
        let target = target.ok_or("transformation without parameter")?;
        let attr_name = target_name.as_deref().map(|name| {
            self.resolve_nameref_text(name)
                .unwrap_or_else(|| name.to_string())
        });
        if all {
            let item = self.fresh("__bash_param_item");
            let out = self.fresh("__bash_param_out");
            let mapped = self.transform_value_expr(ident(&item), &op, None)?;
            let array = array_map_expr(target, &item, &out, mapped);
            return Ok((
                if join_all {
                    array_join(array, ifs_join_sep())
                } else {
                    array
                },
                true,
            ));
        }
        Ok((
            self.transform_value_expr(target, &op, attr_name.as_deref())?,
            false,
        ))
    }

    fn transform_value_expr(
        &mut self,
        target: Expression,
        op: &str,
        attr_name: Option<&str>,
    ) -> R<Expression> {
        if !matches!(
            op,
            "U" | "L" | "u" | "E" | "P" | "Q" | "k" | "K" | "a" | "A"
        ) {
            return Ok(bad_substitution_expr());
        }
        let value = param_value(target.clone());
        Ok(match op {
            "U" => method(value, "toUpperCase", vec![]),
            "L" => method(value, "toLowerCase", vec![]),
            "u" => parts_to_expr(vec![
                Part::Expr(method(
                    method(value.clone(), "charAt", vec![int(0)]),
                    "toUpperCase",
                    vec![],
                )),
                Part::Expr(method(value, "slice", vec![int(1)])),
            ]),
            "E" | "P" => {
                if let Some(name) = attr_name {
                    if let Some(s) = self.variable_values.get(name) {
                        return Ok(lit(&decode_ansi_c(s)));
                    }
                }
                match &target.kind {
                    ExprKind::Lit(Literal::Str(s)) => lit(&decode_ansi_c(s)),
                    _ => value,
                }
            }
            "Q" => {
                if let Some(name) = attr_name {
                    if !self.variable_is_set(name) {
                        return Ok(lit(""));
                    }
                    if let Some(s) = self.variable_values.get(name) {
                        return Ok(lit(&bash_quote(s)));
                    }
                }
                match &target.kind {
                    ExprKind::Lit(Literal::Str(s)) => lit(&bash_quote(s)),
                    _ => bash_quote_expr(value),
                }
            }
            "k" | "K" => {
                if let Some(name) = attr_name {
                    if let Some(entries) = self.array_values.get(name) {
                        let quoted = op == "K";
                        let parts = entries
                            .iter()
                            .flat_map(|(key, value)| {
                                let value = if quoted {
                                    bash_double_quote(value)
                                } else {
                                    value.clone()
                                };
                                [key.clone(), value]
                            })
                            .collect::<Vec<_>>();
                        return Ok(lit(&parts.join(" ")));
                    }
                }
                lit("")
            }
            "a" => lit(&attr_name
                .map(|name| self.attribute_letters(name))
                .unwrap_or_default()),
            "A" => {
                let name = attr_name.unwrap_or_default();
                let attrs = self.attribute_letters(name);
                let flag = if attrs.is_empty() {
                    "--".to_string()
                } else {
                    format!("-{attrs}")
                };
                call_named(
                    "__bash_sprintf",
                    vec![lit("declare %s %s=\"%s\""), lit(&flag), lit(name), value],
                )
            }
            _ => unreachable!(),
        })
    }

    // ── ${ … } ───────────────────────────────────────────────────────────

    fn braced_param_expr(&mut self, pair: Pair<Rule>) -> R<Expression> {
        self.braced_param_expr_with_tilde(pair, true)
    }

    fn braced_param_expr_with_tilde(
        &mut self,
        pair: Pair<Rule>,
        allow_tilde: bool,
    ) -> R<Expression> {
        let inner = pair.into_inner().next().ok_or("empty ${}")?;
        match inner.as_rule() {
            Rule::brace_command_subst_inner => {
                let reply_form = inner.as_str().trim_start().starts_with('|');
                let mut list = None;
                for p in inner.into_inner() {
                    if p.as_rule() == Rule::list {
                        list = Some(p);
                    }
                }
                let body = match list {
                    Some(list) => self.walk_list(list)?,
                    None => Vec::new(),
                };
                if reply_form {
                    Ok(sequence(vec![iife(body), param_value(ident("REPLY"))]))
                } else {
                    Ok(command_substitution_expr(current_shell_capture(
                        body,
                        self.subshell_depth > 0,
                    )))
                }
            }
            Rule::param_base => self.param_base_expr(inner),
            Rule::length_expansion => {
                let base = inner.into_inner().next().ok_or("${#} without parameter")?;
                let base_name = param_base_name(&base);
                let (target, all) = self.param_base_parts(base)?;
                Ok(if all {
                    let assoc_name = base_name
                        .as_ref()
                        .map(|name| {
                            self.resolve_nameref_text(name)
                                .unwrap_or_else(|| name.clone())
                        })
                        .and_then(|name| {
                            name.split_once('[')
                                .map(|(base, _)| base.to_string())
                                .or(Some(name))
                        });
                    if let Some(keys) = assoc_name
                        .as_ref()
                        .and_then(|name| self.static_array_keys(name))
                    {
                        int(keys.len() as i64)
                    } else if assoc_name
                        .as_ref()
                        .is_some_and(|name| self.assoc_arrays.contains(name))
                    {
                        member(call_named("__bash_dict_keys", vec![target]), "length")
                    } else {
                        param_count(target)
                    }
                } else {
                    param_length(target)
                })
            }
            Rule::indirect_subscript_keys => {
                let name = inner
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::name)
                    .map(|p| p.as_str().to_string())
                    .unwrap_or_default();
                let target = self
                    .resolve_nameref_text(&name)
                    .unwrap_or_else(|| name.clone());
                if let Some(keys) = self.static_array_keys(&target) {
                    Ok(lit(&keys.join(" ")))
                } else {
                    Ok(array_join(
                        call_named("__bash_keys", vec![ident(&target)]),
                        lit(" "),
                    ))
                }
            }
            Rule::nameref_prefix_expansion => {
                let sep = if inner.as_str().ends_with('*') {
                    self.static_ifs_first_char()
                        .unwrap_or_else(|| " ".to_string())
                } else {
                    " ".to_string()
                };
                let name = inner
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::name)
                    .map(|p| p.as_str().to_string())
                    .unwrap_or_default();
                Ok(lit(&self.names_with_prefix(&name).join(&sep)))
            }
            Rule::indirect_expansion => {
                let base = inner.into_inner().next().ok_or("${!} without parameter")?;
                if let Some(target) = self.indirect_expr_from_base(&base)? {
                    return Ok(target);
                }
                Ok(lit(""))
            }
            Rule::assign_default
            | Rule::default_value
            | Rule::alternative_value
            | Rule::error_value => {
                let kind = inner.as_rule();
                let mut target = None;
                let mut base_text = String::new();
                let mut op = String::new();
                let mut word = None;
                // `${!ptr:-w}` — the name is read from $ptr first, then the
                // operator applies to what it points at.
                let mut indirect = false;
                for p in inner.into_inner() {
                    match p.as_rule() {
                        Rule::indirect_mark => indirect = true,
                        Rule::param_base => {
                            base_text = p.as_str().to_string();
                            target = if indirect {
                                Some(self.indirect_expr_from_base(&p)?.unwrap_or_else(|| lit("")))
                            } else {
                                Some(self.param_base_parts(p)?.0)
                            }
                        }
                        Rule::assign_default_op
                        | Rule::default_op
                        | Rule::alternative_op
                        | Rule::error_op => op = p.as_str().to_string(),
                        Rule::brace_word => {
                            word = Some(self.brace_word_expr_with_tilde(p, allow_tilde)?)
                        }
                        _ => {}
                    }
                }
                let target = target.ok_or("${} without parameter")?;
                let word = word.unwrap_or_else(|| lit(""));
                // `:` forms treat empty like unset; plain forms test only unset.
                let cond = param_base_test_condition(&base_text, target.clone(), op.starts_with(':'));
                Ok(match kind {
                    Rule::default_value => ternary(cond, target, word),
                    Rule::alternative_value => ternary(cond, word, lit("")),
                    Rule::assign_default => {
                        ternary(cond, target.clone(), assign_expr(target, word))
                    }
                    _ => {
                        let value = self.fresh("__bash_param_value");
                        let cond =
                            param_base_test_condition(&base_text, ident(&value), op.starts_with(':'));
                        iife(vec![
                            let_stmt(&value, target.clone()),
                            Statement::new(StmtKind::If {
                                cond,
                                then_body: vec![Statement::new(StmtKind::Return(Some(ident(
                                    &value,
                                ))))],
                                elifs: Vec::new(),
                                else_body: Some(vec![
                                    bash_stderr_stmt(call_named(
                                        "__bash_sprintf",
                                        vec![lit("%s: %s\n"), lit(&expr_name(&target)), word],
                                    )),
                                    Statement::new(StmtKind::Exit {
                                        status: Some(int(1)),
                                    }),
                                ]),
                            }),
                        ])
                    }
                })
            }
            Rule::substring_expansion => {
                let empty_length = inner.as_str().trim_end().ends_with(':');
                let mut target = None;
                let mut all = false;
                let mut positional_all = false;
                let mut exprs = Vec::new();
                for p in inner.into_inner() {
                    match p.as_rule() {
                        Rule::param_base => {
                            positional_all = matches!(p.as_str(), "@" | "*");
                            let (t, a) = self.param_base_parts(p)?;
                            target = Some(t);
                            all = a;
                        }
                        Rule::arithmetic_expr => exprs.push(self.arith_expr(p)?),
                        _ => {}
                    }
                }
                let target = target.ok_or("substring without parameter")?;
                let mut it = exprs.into_iter();
                let offset = it.next().ok_or("substring without offset")?;
                let length = it.next().or_else(|| empty_length.then(|| int(0)));
                Ok(match (all, length) {
                    // Arrays slice by position; a negative length is an end offset.
                    (true, length) if positional_all => array_join(
                        positional_slice_expr(target, offset, length),
                        ifs_join_sep(),
                    ),
                    (true, length) => array_slice_expr(target, offset, length),
                    (false, length) => self.bash_string_slice_expr(target, offset, length),
                })
            }
            Rule::search_replace => {
                let (expr, _) = self.search_replace_expansion(inner, true)?;
                Ok(expr)
            }
            Rule::prefix_suffix_removal => {
                let (expr, _) = self.prefix_suffix_expansion(inner, true)?;
                Ok(expr)
            }
            Rule::case_modification => {
                let (expr, _) = self.case_mod_expansion(inner, true)?;
                Ok(expr)
            }
            Rule::transform_expansion => {
                let (expr, _) = self.transform_expansion(inner, true)?;
                Ok(expr)
            }
            other => Err(format!(
                "unsupported parameter expansion {other:?}: {}",
                inner.as_str()
            )),
        }
    }

    fn param_base_expr(&mut self, pair: Pair<Rule>) -> R<Expression> {
        if matches!(pair.as_str(), "*" | "@") {
            return Ok(special_param_value(pair.as_str()));
        }
        Ok(param_value(self.param_base_parts(pair)?.0))
    }

    fn names_with_prefix(&self, prefix: &str) -> Vec<String> {
        let mut names: Vec<String> = self
            .variables
            .iter()
            .chain(self.variable_values.keys())
            .chain(self.array_values.keys())
            .filter(|name| name.starts_with(prefix))
            .cloned()
            .collect();
        names.sort();
        names.dedup();
        names
    }

    fn indirect_expr_from_base(&mut self, base: &Pair<Rule>) -> R<Option<Expression>> {
        let Some(name) = param_base_name(base) else {
            return Ok(None);
        };
        if let Some(target) = self.resolve_nameref_text(&name) {
            return Ok(Some(lit(&target)));
        }
        let Some(target) = self.variable_values.get(&name).cloned() else {
            return Ok(None);
        };
        self.dynamic_reference_expr(&target).map(Some)
    }

    fn dynamic_reference_expr(&mut self, target: &str) -> R<Expression> {
        if target.is_empty() {
            return Ok(lit(""));
        }
        if target == "#" || target == "?" || target == "!" || target == "$" || target == "_" {
            return Ok(param_value(special_param(target)));
        }
        if target.chars().all(|c| c.is_ascii_digit()) {
            return Ok(param_value(positional(target)));
        }
        if !is_unset_target(target) {
            return Ok(lit(""));
        }
        if let Some(value) = self.variable_values.get(target) {
            return Ok(lit(value));
        }
        Ok(param_value(self.name_or_element_target(target)?))
    }

    /// The variable reference of a `param_base`, and whether it is `name[@]`/`name[*]`.
    fn param_base_parts(&mut self, pair: Pair<Rule>) -> R<(Expression, bool)> {
        let mut name = None;
        let mut sub = None;
        let mut sub_raw = None;
        let text = pair.as_str().to_string();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::name => name = Some(p.as_str().to_string()),
                Rule::subscript => {
                    sub_raw = Some(p.as_str().to_string());
                    sub = Some(self.subscript_expr(p)?);
                }
                Rule::positional_param => return Ok((positional(p.as_str()), false)),
                Rule::special_param => {
                    return Ok((special_param(p.as_str()), matches!(p.as_str(), "@" | "*")));
                }
                _ => {}
            }
        }
        let mut name = name.ok_or_else(|| format!("bad parameter {text}"))?;
        if name == "FUNCNAME" {
            name = BASH_FUNCNAME.to_string();
        }
        Ok(match sub {
            None => {
                if let Some(target) = self.resolve_nameref_text(&name) {
                    (self.name_or_element_target(&target)?, false)
                } else {
                    (ident(&name), false)
                }
            }
            Some(Subscript::All) => {
                let target = self
                    .resolve_nameref_text(&name)
                    .unwrap_or_else(|| name.clone());
                if target.contains('[') {
                    (self.name_or_element_target(&target)?, false)
                } else {
                    (ident(&target), true)
                }
            }
            Some(Subscript::Key(k)) => {
                let target = self
                    .resolve_nameref_text(&name)
                    .unwrap_or_else(|| name.clone());
                if target.contains('[') {
                    (self.name_or_element_target(&target)?, false)
                } else if self.assoc_arrays.contains(&target) {
                    let key = sub_raw
                        .as_deref()
                        .map(|raw| {
                            raw.strip_prefix('[')
                                .and_then(|s| s.strip_suffix(']'))
                                .unwrap_or(raw)
                        })
                        .map(|raw| self.subscript_assoc_key_expr(raw))
                        .transpose()?
                        .unwrap_or(k);
                    (index(ident(&target), key), false)
                } else {
                    (index(ident(&target), arith_index(k)), false)
                }
            }
        })
    }

    fn subscript_expr(&mut self, pair: Pair<Rule>) -> R<Subscript> {
        let inner = pair.into_inner().next().ok_or("empty subscript")?;
        match inner.as_rule() {
            Rule::subscript_all => Ok(Subscript::All),
            _ => Ok(Subscript::Key(self.subscript_text_expr(inner.as_str())?)),
        }
    }

    /// `arr[text]`: a quoted or `$`-expanded key is a string; a plain name or
    /// arithmetic expression is evaluated (indexed-array style).
    fn subscript_text_expr(&mut self, text: &str) -> R<Expression> {
        let t = text.trim();
        if t.is_empty() {
            return Ok(lit(""));
        }
        if (t.starts_with('"') && t.ends_with('"')) || (t.starts_with('\'') && t.ends_with('\'')) {
            return Ok(lit(&t[1..t.len() - 1]));
        }
        if let Some(n) = parse_bash_integer(t) {
            return Ok(int(n));
        }
        if t.starts_with('$') {
            if let Ok(pairs) = WashmParser::parse(Rule::word, t) {
                for p in pairs {
                    if p.as_rule() == Rule::word {
                        return self.word_expr(p);
                    }
                }
            }
        }
        if let Ok(e) = self.parse_arith_text(t) {
            return Ok(e);
        }
        Ok(lit(t))
    }

    /// Associative-array subscripts are strings after shell expansion; an
    /// unquoted `[user]` is the key `"user"`, not a variable lookup.
    fn subscript_assoc_key_expr(&mut self, text: &str) -> R<Expression> {
        let t = text.trim();
        if t.is_empty() {
            return Ok(lit(""));
        }
        if (t.starts_with('"') && t.ends_with('"')) || (t.starts_with('\'') && t.ends_with('\'')) {
            return Ok(lit(&t[1..t.len() - 1]));
        }
        if t.starts_with('$') {
            if let Ok(pairs) = WashmParser::parse(Rule::word, t) {
                for p in pairs {
                    if p.as_rule() == Rule::word {
                        return self.word_expr(p);
                    }
                }
            }
        }
        Ok(lit(t))
    }

    fn brace_word_expr(&mut self, pair: Pair<Rule>) -> R<Expression> {
        self.brace_word_expr_with_tilde(pair, true)
    }

    fn brace_word_expr_with_tilde(&mut self, pair: Pair<Rule>, allow_tilde: bool) -> R<Expression> {
        let mut parts = Vec::new();
        for p in pair.into_inner() {
            if !allow_tilde && p.as_rule() == Rule::tilde_expansion {
                parts.push(Part::Text(p.as_str().to_string()));
            } else if p.as_rule() == Rule::single_quoted_string {
                parts.push(Part::Text(p.as_str().to_string()));
            } else {
                parts.extend(self.part_to_parts(p)?);
            }
        }
        Ok(parts_to_expr(parts))
    }

    fn last_command_arg(
        &mut self,
        cmd: &Pair<Rule>,
        suffix: &[Pair<Rule>],
    ) -> R<Option<Expression>> {
        for p in suffix.iter().rev() {
            if word_has_runtime_side_effect(p) {
                return Ok(None);
            }
            return Ok(Some(match p.as_rule() {
                Rule::word => self.word_expr(p.clone())?,
                Rule::close_brace_arg => lit("}"),
                Rule::assignment_word => self.assignment_word_as_text(p.clone())?,
                _ => continue,
            }));
        }
        if word_has_runtime_side_effect(cmd) {
            return Ok(None);
        }
        Ok(Some(self.word_expr(cmd.clone())?))
    }
}

// ── helpers ──────────────────────────────────────────────────────────────────

enum AssignTarget {
    Name(String),
    Element(String, Expression),
}

enum Subscript {
    All,
    Key(Expression),
}

fn arith_lvalue_text(pair: &Pair<Rule>) -> Option<String> {
    match pair.as_rule() {
        Rule::arith_variable | Rule::arithmetic_lvalue => {
            arith_assignment_lhs(pair.as_str()).or_else(|| Some(pair.as_str().trim().to_string()))
        }
        Rule::arithmetic_assign => arith_assignment_lhs(pair.as_str()).or_else(|| {
            pair.clone()
                .into_inner()
                .filter_map(|p| arith_lvalue_text(&p))
                .next()
        }),
        Rule::arithmetic_expr
        | Rule::arithmetic_comma
        | Rule::arithmetic_ternary
        | Rule::arithmetic_or
        | Rule::arithmetic_and
        | Rule::arithmetic_bitor
        | Rule::arithmetic_bitxor
        | Rule::arithmetic_bitand
        | Rule::arithmetic_shift
        | Rule::arithmetic_add
        | Rule::arithmetic_mul
        | Rule::arithmetic_eq
        | Rule::arithmetic_cmp
        | Rule::arithmetic_pow
        | Rule::arithmetic_unary
        | Rule::arithmetic_postfix
        | Rule::arith_paren => pair
            .clone()
            .into_inner()
            .filter_map(|p| arith_lvalue_text(&p))
            .next(),
        _ => None,
    }
}

fn arith_assignment_lhs(text: &str) -> Option<String> {
    for op in ["<<=", ">>=", "+=", "-=", "*=", "/=", "%=", "&=", "^=", "|=", "="] {
        if let Some((lhs, _)) = text.split_once(op) {
            let lhs = lhs.trim();
            if is_name(lhs) || (lhs.contains('[') && lhs.ends_with(']')) {
                return Some(lhs.to_string());
            }
        }
    }
    None
}

/// `x+=v`: numbers add, strings concatenate, arrays append.
fn compound_append(target: Expression, value: Expression) -> Statement {
    let next = match &value.kind {
        ExprKind::Array(_) => array_spread(vec![
            (
                ternary(
                    binary(BinOp::StrictEq, target.clone(), undefined()),
                    array(Vec::new()),
                    target.clone(),
                ),
                true,
            ),
            (value, true),
        ]),
        _ => binary(
            BinOp::Add,
            ternary(
                binary(BinOp::StrictEq, target.clone(), undefined()),
                lit(""),
                target.clone(),
            ),
            value,
        ),
    };
    assign_stmt(target, next)
}

#[derive(Clone)]
enum Part {
    Text(String),
    Expr(Expression),
}

fn parts_to_expr(parts: Vec<Part>) -> Expression {
    // Merge adjacent text.
    let mut merged: Vec<Part> = Vec::new();
    for p in parts {
        match (merged.last_mut(), p) {
            (Some(Part::Text(a)), Part::Text(b)) => a.push_str(&b),
            (_, p) => merged.push(p),
        }
    }
    match merged.len() {
        0 => lit(""),
        1 => match merged.pop().unwrap() {
            Part::Text(t) => lit(&t),
            Part::Expr(e) => e,
        },
        _ => Expression::new(ExprKind::Interpolation(
            merged
                .into_iter()
                .map(|p| match p {
                    Part::Text(t) => InterpolPart::Text(t),
                    Part::Expr(e) => InterpolPart::Expr(e),
                })
                .collect(),
        )),
    }
}

fn rewrite_exit_to_return(stmts: &mut [Statement]) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::Exit { status } => {
                stmt.kind = StmtKind::Return(status.take());
            }
            StmtKind::Block(body)
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. }
            | StmtKind::For { body, .. }
            | StmtKind::ForIn { body, .. } => rewrite_exit_to_return(body),
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                rewrite_exit_to_return(then_body);
                for (_, body) in elifs {
                    rewrite_exit_to_return(body);
                }
                if let Some(body) = else_body {
                    rewrite_exit_to_return(body);
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                rewrite_exit_to_return(body);
                for catch in catches {
                    rewrite_exit_to_return(&mut catch.body);
                }
                if let Some(body) = else_body {
                    rewrite_exit_to_return(body);
                }
                if let Some(body) = finally {
                    rewrite_exit_to_return(body);
                }
            }
            _ => {}
        }
    }
}

fn rewrite_exit_to_subshell_throw(stmts: &mut [Statement]) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::Exit { status } => {
                let status = status.take().unwrap_or_else(|| ident("__bash_status"));
                stmt.kind = StmtKind::Throw {
                    expr: Some(object(vec![
                        ("__bash_exit", Expression::bool(true)),
                        ("status", status),
                    ])),
                    cause: None,
                };
            }
            StmtKind::Block(body)
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. }
            | StmtKind::For { body, .. }
            | StmtKind::ForIn { body, .. } => rewrite_exit_to_subshell_throw(body),
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                rewrite_exit_to_subshell_throw(then_body);
                for (_, body) in elifs {
                    rewrite_exit_to_subshell_throw(body);
                }
                if let Some(body) = else_body {
                    rewrite_exit_to_subshell_throw(body);
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                rewrite_exit_to_subshell_throw(body);
                for catch in catches {
                    rewrite_exit_to_subshell_throw(&mut catch.body);
                }
                if let Some(body) = else_body {
                    rewrite_exit_to_subshell_throw(body);
                }
                if let Some(body) = finally {
                    rewrite_exit_to_subshell_throw(body);
                }
            }
            _ => {}
        }
    }
}

fn arith_simple_name(pair: &Pair<Rule>) -> Option<String> {
    if pair.as_rule() != Rule::arith_variable {
        return None;
    }
    let mut name = None;
    for p in pair.clone().into_inner() {
        match p.as_rule() {
            Rule::name => name = Some(p.as_str().to_string()),
            Rule::arith_subscript => return None,
            _ => {}
        }
    }
    name
}

fn arithmetic_expansion_inner(raw: &str) -> Option<&str> {
    raw.strip_prefix("$((")
        .and_then(|s| s.strip_suffix("))"))
        .or_else(|| raw.strip_prefix("$[").and_then(|s| s.strip_suffix(']')))
}

fn arithmetic_cmd_inner(raw: &str) -> Option<&str> {
    raw.trim()
        .strip_prefix("((")
        .and_then(|s| s.strip_suffix("))"))
}

fn expand_assignment_tilde_after_colon(parts: Vec<Part>) -> Vec<Part> {
    let mut out = Vec::new();
    for part in parts {
        match part {
            Part::Text(text) => {
                let mut rest = text.as_str();
                while let Some(pos) = rest.find(":~/") {
                    if pos + 1 > 0 {
                        out.push(Part::Text(rest[..pos + 1].to_string()));
                    }
                    out.push(Part::Expr(param_value(ident("HOME"))));
                    rest = &rest[pos + 2..];
                }
                if !rest.is_empty() {
                    out.push(Part::Text(rest.to_string()));
                }
            }
            other => out.push(other),
        }
    }
    out
}

fn assignment_tilde_literal_expr(text: &str) -> Option<Expression> {
    let mut rest = text.strip_prefix("~/")?;
    let mut parts = vec![
        Part::Expr(param_value(ident("HOME"))),
        Part::Text("/".to_string()),
    ];
    loop {
        let Some(pos) = rest.find(":~/") else {
            if !rest.is_empty() {
                parts.push(Part::Text(rest.to_string()));
            }
            break;
        };
        if pos > 0 {
            parts.push(Part::Text(rest[..pos].to_string()));
        }
        parts.push(Part::Text(":".to_string()));
        parts.push(Part::Expr(param_value(ident("HOME"))));
        parts.push(Part::Text("/".to_string()));
        rest = &rest[pos + 3..];
    }
    Some(parts_to_expr(parts))
}

fn command_modifier_name(pair: &Pair<Rule>) -> Option<String> {
    pair.as_str().split_whitespace().next().map(str::to_string)
}

fn bash_user_variable_predecls(variables: &BTreeSet<String>) -> Vec<VarDeclarator> {
    variables
        .iter()
        .filter(|name| !is_bash_prelude_variable(name))
        .map(|name| declarator(name, Some(undefined())))
        .collect()
}

fn is_bash_prelude_variable(name: &str) -> bool {
    matches!(
        name,
        "__bash_args"
            | "__bash_status"
            | "__bash_ret"
            | "__bash_stdin"
            | "__bash_fds"
            | "__bash_stdout_null"
            | "__bash_stdout_to_stderr"
            | "__bash_stderr_to_stdout"
            | "__bash_stderr_null"
            | "__bash_abort_line"
            | "__bash_line"
            | "BASH_REMATCH"
            | "PIPESTATUS"
            | "__bash_funcname"
            | "__bash_had"
            | "__bash_fields"
            | "__bash_last_arg"
            | "__bash_pid"
            | "BASH_SUBSHELL"
            | "__bash_last_bg_pid"
            | "__bash_job_seq"
            | "__bash_jobs"
            | "__bash_coprocs"
            | "__bash_flags"
            | "TIMEFORMAT"
            | "HOME"
            | "PWD"
            | "OLDPWD"
    )
}

fn function_source_has_local_dash(source: &str) -> bool {
    source.split(['\n', ';']).any(|segment| {
        let trimmed = segment.trim_start().trim_start_matches('{').trim_start();
        trimmed == "local -" || trimmed.starts_with("local - ")
    })
}

fn collect_function_scoped_names(stmts: &[Statement], names: &mut Vec<String>) {
    for stmt in stmts {
        match &stmt.kind {
            StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::FunctionScoped,
            } => {
                for decl in declarations {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if !names.iter().any(|existing| existing == name) {
                            names.push(name.clone());
                        }
                    }
                }
            }
            StmtKind::Block(body)
            | StmtKind::FunctionDecl { body, .. }
            | StmtKind::For { body, .. }
            | StmtKind::ForIn { body, .. }
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. }
            | StmtKind::With { body, .. }
            | StmtKind::Using { body, .. }
            | StmtKind::Lock { body, .. } => collect_function_scoped_names(body, names),
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                collect_function_scoped_names(then_body, names);
                for (_, body) in elifs {
                    collect_function_scoped_names(body, names);
                }
                if let Some(body) = else_body {
                    collect_function_scoped_names(body, names);
                }
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    collect_function_scoped_names(&case.body, names);
                }
                if let Some(body) = default {
                    collect_function_scoped_names(body, names);
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                collect_function_scoped_names(body, names);
                for catch in catches {
                    collect_function_scoped_names(&catch.body, names);
                }
                if let Some(body) = else_body {
                    collect_function_scoped_names(body, names);
                }
                if let Some(body) = finally {
                    collect_function_scoped_names(body, names);
                }
            }
            _ => {}
        }
    }
}

fn rewrite_function_scoped_decls(stmts: &mut Vec<Statement>) {
    let mut rewritten = Vec::with_capacity(stmts.len());
    for mut stmt in stmts.drain(..) {
        match &mut stmt.kind {
            StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::FunctionScoped,
            } => {
                for decl in declarations.drain(..) {
                    if let BindingPattern::Ident(name) = decl.pattern {
                        rewritten.push(assign_stmt(
                            ident(&name),
                            decl.init.unwrap_or_else(undefined),
                        ));
                    }
                }
            }
            StmtKind::Block(body)
            | StmtKind::FunctionDecl { body, .. }
            | StmtKind::For { body, .. }
            | StmtKind::ForIn { body, .. }
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. }
            | StmtKind::With { body, .. }
            | StmtKind::Using { body, .. }
            | StmtKind::Lock { body, .. } => {
                rewrite_function_scoped_decls(body);
                rewritten.push(stmt);
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                rewrite_function_scoped_decls(then_body);
                for (_, body) in elifs {
                    rewrite_function_scoped_decls(body);
                }
                if let Some(body) = else_body {
                    rewrite_function_scoped_decls(body);
                }
                rewritten.push(stmt);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    rewrite_function_scoped_decls(&mut case.body);
                }
                if let Some(body) = default {
                    rewrite_function_scoped_decls(body);
                }
                rewritten.push(stmt);
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                rewrite_function_scoped_decls(body);
                for catch in catches {
                    rewrite_function_scoped_decls(&mut catch.body);
                }
                if let Some(body) = else_body {
                    rewrite_function_scoped_decls(body);
                }
                if let Some(body) = finally {
                    rewrite_function_scoped_decls(body);
                }
                rewritten.push(stmt);
            }
            _ => rewritten.push(stmt),
        }
    }
    *stmts = rewritten;
}

fn simple_command_literal_parts(pair: Pair<Rule>) -> Option<(String, Vec<Pair<Rule>>)> {
    let mut prefixes = false;
    let mut cmd_word: Option<Pair<Rule>> = None;
    let mut suffix = Vec::new();
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::cmd_prefix => prefixes = true,
            Rule::cmd_modifier => return None,
            Rule::cmd_word => cmd_word = inner.into_inner().next(),
            Rule::cmd_suffix => {
                for item in inner.into_inner() {
                    if item.as_rule() == Rule::redirection {
                        return None;
                    }
                    suffix.push(item);
                }
            }
            Rule::redirection => return None,
            _ => {}
        }
    }
    if prefixes {
        return None;
    }
    let word = cmd_word?;
    Some((literal_word_text(&word)?, suffix))
}

fn assignment_word_has_command_substitution(pair: &Pair<Rule>) -> bool {
    pair.clone().into_inner().any(pair_has_command_substitution)
}

fn assignment_word_target_text(pair: &Pair<Rule>) -> Option<String> {
    if pair.as_rule() != Rule::assignment_word {
        return None;
    }
    pair.clone()
        .into_inner()
        .find(|p| p.as_rule() == Rule::assign_target)
        .map(|p| p.as_str().to_string())
}

fn pair_has_command_substitution(pair: Pair<Rule>) -> bool {
    matches!(
        pair.as_rule(),
        Rule::dollar_paren_subst | Rule::backtick_subst | Rule::brace_command_subst_inner
    ) || pair.into_inner().any(pair_has_command_substitution)
}

fn is_single_bracket_arith_form(words: &[Pair<Rule>]) -> bool {
    words.len() == 3
        && matches!(
            words[1].as_str(),
            "-eq" | "-ne" | "-lt" | "-le" | "-gt" | "-ge"
        )
}

fn is_test_unary_operator(op: &str) -> bool {
    matches!(
        op,
        "-a" | "-b"
            | "-c"
            | "-d"
            | "-e"
            | "-f"
            | "-g"
            | "-h"
            | "-k"
            | "-L"
            | "-N"
            | "-O"
            | "-G"
            | "-p"
            | "-r"
            | "-s"
            | "-S"
            | "-t"
            | "-u"
            | "-w"
            | "-x"
            | "-n"
            | "-z"
            | "-v"
            | "-R"
            | "-o"
    )
}

fn double_bracket_has_unary_without_operand(pair: &Pair<Rule>) -> bool {
    let raw = pair.as_str().trim();
    let Some(inner) = raw.strip_prefix("[[").and_then(|s| s.strip_suffix("]]")) else {
        return false;
    };
    let tokens = inner.split_whitespace().collect::<Vec<_>>();
    match tokens.as_slice() {
        [op] => is_test_unary_operator(op),
        ["!", op] => is_test_unary_operator(op),
        _ => false,
    }
}

fn double_bracket_uses_legacy_boolean_op(pair: &Pair<Rule>) -> bool {
    pair.clone().into_inner().any(|p| match p.as_rule() {
        Rule::test_and_op => p.as_str() == "-a",
        Rule::test_or_op => p.as_str() == "-o",
        _ => double_bracket_uses_legacy_boolean_op(&p),
    })
}

fn single_bracket_literal_arith_args(pair: &Pair<Rule>) -> Option<Vec<Expression>> {
    let raw = pair.as_str().trim();
    let inner = raw.strip_prefix('[')?.strip_suffix(']')?.trim();
    let words = inner.split_whitespace().collect::<Vec<_>>();
    if words.len() != 3 || !matches!(words[1], "-eq" | "-ne" | "-lt" | "-le" | "-gt" | "-ge") {
        return None;
    }
    if words
        .iter()
        .any(|w| w.contains('$') || w.contains('"') || w.contains('\'') || w.contains('\\'))
    {
        return None;
    }
    Some(words.into_iter().map(lit).collect())
}

fn single_bracket_arith_error_status_expr(words: &[Pair<Rule>], command: &str) -> Option<Expression> {
    if words.len() != 3 {
        return None;
    }
    let op = literal_word_text(&words[1])?;
    if !matches!(op.as_str(), "-eq" | "-ne" | "-lt" | "-le" | "-gt" | "-ge") {
        return None;
    }
    for operand in [&words[0], &words[2]] {
        let text = literal_word_text(operand)?;
        if parse_bash_integer(&text).is_none() {
            return Some(sequence(vec![
                bash_stderr_write_expr(lit(&format!(
                    "bash: {command}: {text}: integer expected\n"
                ))),
                int(2),
            ]));
        }
    }
    None
}

fn single_bracket_raw_arith_error_status_expr(raw: &str, command: &str) -> Option<Expression> {
    let inner = raw.trim().strip_prefix('[')?.strip_suffix(']')?.trim();
    let tokens = shell_word_sources(inner)?;
    if tokens.len() != 3 {
        return None;
    }
    if !matches!(tokens[1].as_str(), "-eq" | "-ne" | "-lt" | "-le" | "-gt" | "-ge") {
        return None;
    }
    for token in [&tokens[0], &tokens[2]] {
        let mut pairs = WashmParser::parse(Rule::word, token).ok()?;
        let word = pairs.next()?;
        let text = literal_word_text(&word)?;
        if parse_bash_integer(&text).is_none() {
            return Some(sequence(vec![
                bash_stderr_write_expr(lit(&format!(
                    "bash: {command}: {text}: integer expected\n"
                ))),
                int(2),
            ]));
        }
    }
    None
}

fn single_bracket_pair_arith_error_status_expr(
    pair: &Pair<Rule>,
    command: &str,
) -> Option<Expression> {
    let arith = find_test_binary_arith(pair)?;
    let words = arith
        .clone()
        .into_inner()
        .filter(|p| p.as_rule() == Rule::word)
        .collect::<Vec<_>>();
    single_bracket_arith_error_status_expr(&words, command)
}

fn find_test_binary_arith<'i>(pair: &Pair<'i, Rule>) -> Option<Pair<'i, Rule>> {
    if pair.as_rule() == Rule::test_binary_arith {
        return Some(pair.clone());
    }
    pair.clone()
        .into_inner()
        .find_map(|p| find_test_binary_arith(&p))
}

fn here_string_raw_fd(raw: &str) -> Option<String> {
    let (fd, _) = raw.split_once("<<<")?;
    (!fd.is_empty() && fd.chars().all(|c| c.is_ascii_digit())).then(|| fd.to_string())
}

fn redirection_is_stdin_form(pair: &Pair<Rule>) -> bool {
    let raw = pair.as_str();
    raw.contains("<<<") || raw.contains("<<") || raw.contains('<')
}

fn is_name(s: &str) -> bool {
    let mut chars = s.chars();
    match chars.next() {
        Some(c) if c.is_ascii_alphabetic() || c == '_' => {
            chars.all(|c| c.is_ascii_alphanumeric() || c == '_')
        }
        _ => false,
    }
}

fn is_unset_target(s: &str) -> bool {
    if let Some((name, rest)) = s.split_once('[') {
        return is_name(name) && rest.ends_with(']');
    }
    is_name(s)
}

fn is_valid_indirect_target(s: &str) -> bool {
    if is_unset_target(s) {
        return true;
    }
    matches!(s, "#" | "$" | "!" | "?" | "-" | "_" | "@" | "*")
        || s.chars().all(|c| c.is_ascii_digit())
}

fn is_readonly_parameter_target(s: &str) -> bool {
    let base = s.split_once('[').map(|(name, _)| name).unwrap_or(s);
    matches!(base, "#" | "@" | "*" | "?" | "$" | "!" | "-" | "_" | "0")
        || (!base.is_empty() && base.chars().all(|c| c.is_ascii_digit()))
}

fn eval_obvious_syntax_error(src: &str) -> Option<String> {
    let trimmed = src.trim_start();
    if has_unterminated_quote(trimmed) {
        return Some("unexpected EOF while looking for matching quote".to_string());
    }
    if trimmed.starts_with(';') || trimmed.starts_with("&&") || trimmed.starts_with('|') {
        return Some("syntax error near unexpected token".to_string());
    }
    if has_unclosed_double_bracket(trimmed) {
        return Some("syntax error".to_string());
    }
    if let Some(rest) = trimmed.strip_prefix('{') {
        if rest.chars().next().is_some_and(|c| !c.is_whitespace()) {
            return Some("syntax error near unexpected token `{`".to_string());
        }
    }
    if is_empty_subshell_text(trimmed) {
        return Some("syntax error near unexpected token `)'".to_string());
    }
    if has_unmatched_closing_paren(trimmed) {
        return Some("syntax error near unexpected token `)'".to_string());
    }
    if has_dangling_redirection(trimmed) {
        return Some("syntax error near unexpected token `newline'".to_string());
    }
    if has_simple_command_function_body(trimmed) {
        return Some("syntax error near unexpected token".to_string());
    }
    if has_keyword_function_definition(trimmed) {
        return Some("syntax error".to_string());
    }
    if trimmed.starts_with("for ")
        && trimmed.contains("; done")
        && !trimmed.contains(" do")
        && !trimmed.contains("; do")
    {
        return Some("syntax error near unexpected token `done'".to_string());
    }
    if trimmed.starts_with("if ")
        && trimmed.contains("; fi")
        && !trimmed.contains(" then")
        && !trimmed.contains("; then")
    {
        return Some("syntax error near unexpected token `fi'".to_string());
    }
    if trimmed.contains("&;") {
        return Some("syntax error near unexpected token `;'".to_string());
    }
    if trimmed.contains(";;") && !trimmed.starts_with("case ") {
        return Some("syntax error near unexpected token `;;'".to_string());
    }
    None
}

fn bash_c_syntax_error_expr(source: &str, message: &str) -> Expression {
    let line = syntax_error_line(source).unwrap_or(1);
    sequence(vec![
        bash_stderr_write_expr(lit(&format!("bash: line {line}: {message}\n"))),
        int(2),
    ])
}

fn syntax_error_line(src: &str) -> Option<usize> {
    let mut depth = 0i32;
    let mut line = 1usize;
    let mut chars = src.chars().peekable();
    let mut in_single = false;
    let mut in_double = false;
    while let Some(ch) = chars.next() {
        match ch {
            '\n' => line += 1,
            '\\' if !in_single => {
                if matches!(chars.peek(), Some('\n')) {
                    line += 1;
                }
                chars.next();
            }
            '\'' if !in_double => in_single = !in_single,
            '"' if !in_single => in_double = !in_double,
            '(' if !in_single && !in_double => depth += 1,
            ')' if !in_single && !in_double => {
                if depth == 0 {
                    return Some(line);
                }
                depth -= 1;
            }
            _ => {}
        }
    }
    Some(line)
}

fn has_unterminated_quote(src: &str) -> bool {
    let mut chars = src.chars().peekable();
    let mut in_single = false;
    let mut in_double = false;
    while let Some(ch) = chars.next() {
        match ch {
            '\\' if !in_single => {
                chars.next();
            }
            '\'' if !in_double => in_single = !in_single,
            '"' if !in_single => in_double = !in_double,
            _ => {}
        }
    }
    in_single || in_double
}

fn has_unclosed_double_bracket(src: &str) -> bool {
    let mut opens = 0usize;
    let mut chars = src.chars().peekable();
    let mut in_single = false;
    let mut in_double = false;
    while let Some(ch) = chars.next() {
        match ch {
            '\\' if !in_single => {
                chars.next();
            }
            '\'' if !in_double => in_single = !in_single,
            '"' if !in_single => in_double = !in_double,
            '[' if !in_single && !in_double && matches!(chars.peek(), Some('[')) => {
                chars.next();
                opens += 1;
            }
            ']' if !in_single && !in_double && matches!(chars.peek(), Some(']')) => {
                chars.next();
                opens = opens.saturating_sub(1);
            }
            _ => {}
        }
    }
    opens > 0
}

fn eval_obvious_expansion_error(
    src: &str,
    variable_values: &HashMap<String, String>,
) -> Option<&'static str> {
    if src.contains("${${") || src.contains("${}") {
        return Some("bash: bad substitution");
    }
    if src.contains("@Q@") || src.contains("@Z}") {
        return Some("bash: bad substitution");
    }
    if let Some(message) = eval_obvious_arithmetic_error(src, variable_values) {
        return Some(message);
    }
    if has_negative_substring_length_error(src, variable_values) {
        return Some("bash: substring expression < 0");
    }
    None
}

fn eval_obvious_arithmetic_error(
    src: &str,
    variable_values: &HashMap<String, String>,
) -> Option<&'static str> {
    for expr in arithmetic_expansion_texts(src) {
        if arithmetic_has_self_reference(expr.trim(), variable_values) {
            return Some("bash: expression recursion level exceeded");
        }
        let text = arithmetic_obvious_value(expr.trim(), variable_values);
        let text = text.trim();
        if arithmetic_starts_with_non_lvalue_assignment(text)
            || arithmetic_starts_with_non_lvalue_increment(text)
        {
            return Some("bash: attempted assignment to non-variable");
        }
        if arithmetic_has_float_literal(text) {
            return Some("bash: arithmetic syntax error");
        }
        if arithmetic_has_pow_assignment(text) {
            return Some("bash: operand expected");
        }
        if arithmetic_has_negative_exponent(text) {
            return Some("bash: exponent less than 0");
        }
        if arithmetic_has_invalid_base(text) {
            return Some("bash: invalid arithmetic base");
        }
        if arithmetic_has_digit_out_of_range(text) {
            return Some("bash: value too great for base");
        }
        if arithmetic_has_zero_divisor(text) {
            return Some("bash: division by 0");
        }
        if arithmetic_has_missing_operand(text) {
            return Some("bash: operand expected");
        }
    }
    None
}

fn bash_arithmetic_error_message(text: &str) -> Option<&'static str> {
    let text = text.trim();
    if arithmetic_has_float_literal(text) {
        return Some("arithmetic syntax error");
    }
    if arithmetic_has_pow_assignment(text) {
        return Some("operand expected");
    }
    if arithmetic_has_negative_exponent(text) {
        return Some("exponent less than 0");
    }
    if arithmetic_has_invalid_base(text) {
        return Some("invalid arithmetic base");
    }
    if arithmetic_has_digit_out_of_range(text) {
        return Some("value too great for base");
    }
    if arithmetic_has_zero_divisor(text) {
        return Some("division by 0");
    }
    if arithmetic_has_missing_operand(text) {
        return Some("operand expected");
    }
    None
}

fn bash_arithmetic_error_expr(message: &str) -> Expression {
    iife(vec![
        bash_stderr_stmt(lit(&format!("bash: {message}\n"))),
        assign_stmt(ident("__bash_status"), int(1)),
        Statement::new(StmtKind::Return(Some(lit("")))),
    ])
}

fn arithmetic_has_self_reference(text: &str, variable_values: &HashMap<String, String>) -> bool {
    let mut name = text.trim();
    let mut seen = HashSet::new();
    while is_name(name) {
        if !seen.insert(name.to_string()) {
            return true;
        }
        let Some(value) = variable_values.get(name) else {
            return false;
        };
        name = value.trim();
    }
    false
}

fn arithmetic_obvious_value<'a>(
    text: &'a str,
    variable_values: &'a HashMap<String, String>,
) -> &'a str {
    let trimmed = text.trim();
    if is_name(trimmed) {
        variable_values
            .get(trimmed)
            .map(|s| s.as_str())
            .unwrap_or(trimmed)
    } else {
        trimmed
    }
}

fn arithmetic_expansion_texts(src: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut rest = src;
    loop {
        let paren_pos = rest.find("$((");
        let bracket_pos = rest.find("$[");
        let Some((pos, prefix_len, suffix)) = (match (paren_pos, bracket_pos) {
            (Some(a), Some(b)) if a <= b => Some((a, 3, "))")),
            (Some(_), Some(b)) => Some((b, 2, "]")),
            (Some(a), None) => Some((a, 3, "))")),
            (None, Some(b)) => Some((b, 2, "]")),
            (None, None) => None,
        }) else {
            break;
        };
        let after = &rest[pos + prefix_len..];
        let Some(end) = after.find(suffix) else {
            break;
        };
        out.push(&after[..end]);
        rest = &after[end + suffix.len()..];
    }
    out
}

fn arithmetic_starts_with_non_lvalue_assignment(text: &str) -> bool {
    let mut chars = text.trim_start().chars().peekable();
    if !chars.next().is_some_and(|c| c.is_ascii_digit()) {
        return false;
    }
    while chars.peek().is_some_and(|c| c.is_ascii_digit()) {
        chars.next();
    }
    while chars.peek().is_some_and(|c| c.is_whitespace()) {
        chars.next();
    }
    chars.next() == Some('=') && chars.peek() != Some(&'=')
}

fn arithmetic_starts_with_non_lvalue_increment(text: &str) -> bool {
    let trimmed = text.trim_start();
    let digit_count = trimmed.chars().take_while(|c| c.is_ascii_digit()).count();
    digit_count > 0 && trimmed[digit_count..].trim_start().starts_with("++")
}

fn arithmetic_has_float_literal(text: &str) -> bool {
    let bytes = text.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if !bytes[i].is_ascii_digit() {
            i += 1;
            continue;
        }
        let start = i;
        while i < bytes.len() && bytes[i].is_ascii_digit() {
            i += 1;
        }
        if i < bytes.len()
            && bytes[i] == b'.'
            && bytes.get(i + 1).is_some_and(|c| c.is_ascii_digit())
        {
            let prev = start.checked_sub(1).and_then(|idx| bytes.get(idx)).copied();
            if prev.is_none_or(|c| !c.is_ascii_alphanumeric() && c != b'_' && c != b'#') {
                return true;
            }
        }
    }
    false
}

fn arithmetic_has_pow_assignment(text: &str) -> bool {
    text.contains("**=")
}

fn arithmetic_has_negative_exponent(text: &str) -> bool {
    let mut rest = text;
    while let Some(pos) = rest.find("**") {
        let after = rest[pos + 2..].trim_start();
        if after.starts_with('-') {
            return true;
        }
        rest = &after[after
            .char_indices()
            .nth(1)
            .map(|(i, _)| i)
            .unwrap_or(after.len())..];
    }
    false
}

fn arithmetic_has_invalid_base(text: &str) -> bool {
    arithmetic_tokens(text).into_iter().any(|token| {
        token
            .split_once('#')
            .and_then(|(base, digits)| {
                (!digits.is_empty())
                    .then(|| base.trim_start_matches(['+', '-']).parse::<u32>().ok())
                    .flatten()
            })
            .is_some_and(|base| !(2..=64).contains(&base))
    })
}

fn arithmetic_has_digit_out_of_range(text: &str) -> bool {
    arithmetic_tokens(text).into_iter().any(|token| {
        bash_integer_error(token)
            .is_some_and(|err| matches!(err, BashIntegerError::DigitOutOfRange))
    })
}

fn arithmetic_has_zero_divisor(text: &str) -> bool {
    let bytes = text.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] != b'/' && bytes[i] != b'%' {
            i += 1;
            continue;
        }
        i += 1;
        while i < bytes.len() && bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        if i < bytes.len() && bytes[i] == b'0' {
            let next = bytes.get(i + 1).copied();
            if next.is_none_or(|c| !c.is_ascii_alphanumeric() && c != b'_') {
                return true;
            }
        }
    }
    false
}

fn arithmetic_has_missing_operand(text: &str) -> bool {
    let trimmed = text.trim_end();
    if trimmed.ends_with("++") || trimmed.ends_with("--") {
        return false;
    }
    trimmed.ends_with('+')
        || trimmed.ends_with('-')
        || trimmed.ends_with('*')
        || trimmed.ends_with('/')
        || trimmed.ends_with('%')
        || trimmed.ends_with("<<")
        || trimmed.ends_with(">>")
        || trimmed.ends_with('&')
        || trimmed.ends_with('|')
        || trimmed.ends_with('^')
}

fn has_negative_substring_length_error(
    src: &str,
    variable_values: &HashMap<String, String>,
) -> bool {
    let mut rest = src;
    while let Some(pos) = rest.find("${") {
        rest = &rest[pos + 2..];
        let Some(end) = rest.find('}') else {
            break;
        };
        let inner = &rest[..end];
        rest = &rest[end + 1..];
        let mut parts = inner.split(':');
        let Some(name) = parts.next() else {
            continue;
        };
        if name.is_empty() || name == "-" {
            continue;
        }
        let Some(offset) = parts.next().and_then(|s| s.trim().parse::<i64>().ok()) else {
            continue;
        };
        let Some(length) = parts.next().and_then(|s| s.trim().parse::<i64>().ok()) else {
            continue;
        };
        if length < 0 {
            let value_len = variable_values
                .get(name)
                .map(|value| value.chars().count() as i64)
                .unwrap_or(0);
            if value_len + length < offset {
                return true;
            }
        }
    }
    false
}

fn is_empty_subshell_text(src: &str) -> bool {
    src.strip_prefix('(')
        .and_then(|s| s.strip_suffix(')'))
        .is_some_and(|inner| inner.trim().is_empty())
}

fn has_dangling_redirection(src: &str) -> bool {
    let trimmed = src.trim_end();
    [">", ">>", ">|", "<", "<>", ">&", "<&", "&>", "&>>"]
        .iter()
        .any(|op| trimmed.ends_with(op))
}

fn has_simple_command_function_body(src: &str) -> bool {
    let Some((_, rest)) = src.split_once("()") else {
        return false;
    };
    let body = rest.trim_start();
    !body.is_empty()
        && !body.starts_with('{')
        && !body.starts_with('(')
        && !body.starts_with("[[")
        && !body.starts_with("if ")
        && !body.starts_with("for ")
        && !body.starts_with("while ")
        && !body.starts_with("until ")
        && !body.starts_with("case ")
}

fn has_unmatched_closing_paren(src: &str) -> bool {
    let mut depth = 0i32;
    let mut chars = src.chars().peekable();
    let mut in_single = false;
    let mut in_double = false;
    while let Some(ch) = chars.next() {
        match ch {
            '\\' if !in_single => {
                chars.next();
            }
            '\'' if !in_double => in_single = !in_single,
            '"' if !in_single => in_double = !in_double,
            '(' if !in_single && !in_double => depth += 1,
            ')' if !in_single && !in_double => {
                if depth == 0 {
                    return true;
                }
                depth -= 1;
            }
            _ => {}
        }
    }
    false
}

fn expr_name(e: &Expression) -> String {
    match &e.kind {
        ExprKind::Ident(n) => n.clone(),
        _ => String::new(),
    }
}

fn param_base_name(pair: &Pair<Rule>) -> Option<String> {
    if pair.as_rule() != Rule::param_base {
        return None;
    }
    pair.clone()
        .into_inner()
        .find(|p| {
            matches!(
                p.as_rule(),
                Rule::name | Rule::positional_param | Rule::special_param
            )
        })
        .map(|p| p.as_str().to_string())
}

fn parse_positional_slice_text(text: &str) -> Option<(Expression, Option<Expression>)> {
    let mut parts = text.splitn(2, ':');
    let offset = parts.next()?.trim().parse::<i64>().ok()?;
    let length = match parts.next() {
        Some(rest) if rest.trim().is_empty() => Some(int(0)),
        Some(rest) => Some(int(rest.trim().parse::<i64>().ok()?)),
        None => None,
    };
    Some((int(offset), length))
}

fn static_transform_braced(raw: &str) -> Option<(&str, &str)> {
    let inner = raw.strip_prefix("${")?.strip_suffix('}')?;
    let (name, op) = inner.split_once('@')?;
    if !is_name(name) || op.len() != 1 {
        return None;
    }
    Some((name, op))
}

fn expr_is_negative_literal(e: &Expression) -> bool {
    matches!(&e.kind, ExprKind::Lit(Literal::Int(n)) if *n < 0)
}

/// `__bash_procsub("<", () => …)`.
fn is_procsub(e: &Expression) -> bool {
    matches!(&e.kind, ExprKind::Call { callee, .. } if matches!(&callee.kind, ExprKind::Ident(n) if n == "__bash_procsub"))
}

fn procsub_body(e: Expression) -> Vec<Statement> {
    if let ExprKind::Call { args, .. } = e.kind {
        if let Some(body) = args.into_iter().nth(1) {
            if let ExprKind::Lambda {
                body: LambdaBody::Block(stmts),
                ..
            } = body.value.kind
            {
                return stmts;
            }
        }
    }
    Vec::new()
}

/// The output of `<(…)` used as a redirection source.
fn procsub_capture(e: Expression) -> Expression {
    let body = procsub_body(e);
    if body.is_empty() {
        lit("")
    } else {
        capture(body)
    }
}

/// `$1` … `$9` and `${10}`: the function's (or script's) argument list.
fn positional(digits: &str) -> Expression {
    let n: i64 = digits.parse().unwrap_or(0);
    if n == 0 {
        return lit("bash");
    }
    index(ident(BASH_ARGS), int(n - 1))
}

fn special_param(s: &str) -> Expression {
    match s {
        "#" => member(ident(BASH_ARGS), "length"),
        "@" | "*" => ident(BASH_ARGS),
        "?" => ident("__bash_status"),
        "$" => ident("__bash_pid"),
        "!" => ident("__bash_last_bg_pid"),
        "-" => ident("__bash_flags"),
        _ => ident("__bash_last_arg"),
    }
}

fn param_base_test_condition(base: &str, value: Expression, empty_counts_unset: bool) -> Expression {
    if base == "#" {
        return if empty_counts_unset {
            binary(BinOp::Gt, member(ident(BASH_ARGS), "length"), int(0))
        } else {
            Expression::bool(true)
        };
    }
    if matches!(base, "?" | "$" | "-" | "_") {
        return Expression::bool(true);
    }
    if base == "!" {
        return if empty_counts_unset {
            binary(BinOp::NotEq, ident("__bash_last_bg_pid"), int(0))
        } else {
            Expression::bool(true)
        };
    }
    if base == "@" || base == "*" {
        return if empty_counts_unset {
            binary(BinOp::Gt, member(ident(BASH_ARGS), "length"), int(0))
        } else {
            Expression::bool(true)
        };
    }
    if base.chars().all(|c| c.is_ascii_digit()) {
        return if empty_counts_unset {
            binary(
                BinOp::And,
                binary(BinOp::StrictNotEq, value.clone(), undefined()),
                binary(BinOp::StrictNotEq, param_value(value), lit("")),
            )
        } else {
            binary(BinOp::StrictNotEq, value, undefined())
        };
    }
    if empty_counts_unset {
        value
    } else {
        binary(BinOp::StrictNotEq, value, undefined())
    }
}

fn special_param_value(s: &str) -> Expression {
    match s {
        "@" | "*" => array_join(ident(BASH_ARGS), ifs_join_sep()),
        _ => param_value(special_param(s)),
    }
}

/// The text of a word made only of plain characters, else `None`.
fn literal_word_text(pair: &Pair<Rule>) -> Option<String> {
    if pair.as_rule() != Rule::word {
        return None;
    }
    let mut out = String::new();
    for p in pair.clone().into_inner() {
        match p.as_rule() {
            Rule::bare_word => out.push_str(&unescape_bare(p.as_str())),
            Rule::close_brace_tail => out.push('}'),
            Rule::single_quoted_string => {
                let s = p.as_str();
                out.push_str(&s[1..s.len() - 1]);
            }
            Rule::quoted_string => {
                for q in p.into_inner() {
                    match q.as_rule() {
                        Rule::dq_text => out.push_str(q.as_str()),
                        Rule::dq_escape => out.push_str(&q.as_str()[1..]),
                        _ => return None,
                    }
                }
            }
            _ => return None,
        }
    }
    Some(out)
}

fn exact_braced_param_word<'i>(pair: &Pair<'i, Rule>) -> Option<Pair<'i, Rule>> {
    if pair.as_rule() != Rule::word {
        return None;
    }
    let mut parts: Vec<Pair<Rule>> = pair.clone().into_inner().collect();
    if parts.len() == 1 && parts[0].as_rule() == Rule::quoted_string {
        parts = parts.remove(0).into_inner().collect();
    }
    if parts.len() == 1 && parts[0].as_rule() == Rule::braced_param {
        return Some(parts.remove(0));
    }
    None
}

fn replace_static_arith_expansions(input: &str) -> Option<String> {
    let mut out = String::new();
    let mut rest = input;
    while let Some(start) = rest.find("$((") {
        out.push_str(&rest[..start]);
        let after = &rest[start + 3..];
        let end = after.find("))")?;
        let expr = &after[..end];
        out.push_str(&eval_const_arith(expr)?.to_string());
        rest = &after[end + 2..];
    }
    out.push_str(rest);
    Some(out)
}

fn replace_static_command_substitutions(input: &str) -> Option<String> {
    let mut out = String::new();
    let mut rest = input;
    while let Some(start) = rest.find("$(") {
        out.push_str(&rest[..start]);
        let after = &rest[start + 2..];
        let end = after.find(')')?;
        let command = after[..end].trim();
        out.push_str(&static_command_output(command)?);
        rest = &after[end + 1..];
    }
    out.push_str(rest);
    Some(out)
}

fn static_command_output(command: &str) -> Option<String> {
    if raw_may_contain_brace_expansion(command) {
        return None;
    }
    if let Some(rest) = command.strip_prefix("printf ") {
        let mut parts = shell_static_words(rest)?;
        let fmt = parts
            .is_empty()
            .then(String::new)
            .unwrap_or_else(|| parts.remove(0));
        let text = decode_printf_escapes(&fmt);
        return Some(text.trim_end_matches('\n').to_string());
    }
    if let Some(rest) = command.strip_prefix("echo ") {
        return Some(shell_static_words(rest)?.join(" "));
    }
    None
}

fn shell_static_words(mut input: &str) -> Option<Vec<String>> {
    let mut out = Vec::new();
    while !input.trim_start().is_empty() {
        input = input.trim_start();
        let first = input.chars().next()?;
        if first == '\'' || first == '"' {
            let mut escaped = false;
            let mut end = None;
            for (idx, ch) in input.char_indices().skip(1) {
                if first == '"' && ch == '\\' && !escaped {
                    escaped = true;
                    continue;
                }
                if ch == first && !escaped {
                    end = Some(idx);
                    break;
                }
                escaped = false;
            }
            let end = end?;
            out.push(input[1..end].to_string());
            input = &input[end + 1..];
        } else {
            let end = input
                .char_indices()
                .find(|(_, ch)| ch.is_whitespace())
                .map(|(idx, _)| idx)
                .unwrap_or(input.len());
            out.push(unescape_bare(&input[..end]));
            input = &input[end..];
        }
    }
    Some(out)
}

fn shell_word_sources(mut input: &str) -> Option<Vec<String>> {
    let mut out = Vec::new();
    while !input.trim_start().is_empty() {
        input = input.trim_start();
        let mut escaped = false;
        let mut quote: Option<char> = None;
        let mut end = input.len();
        for (idx, ch) in input.char_indices() {
            if escaped {
                escaped = false;
                continue;
            }
            if quote == Some('"') && ch == '\\' {
                escaped = true;
                continue;
            }
            if quote.is_none() && ch == '\\' {
                escaped = true;
                continue;
            }
            if let Some(q) = quote {
                if ch == q {
                    quote = None;
                }
                continue;
            }
            if ch == '\'' || ch == '"' {
                quote = Some(ch);
                continue;
            }
            if ch.is_whitespace() {
                end = idx;
                break;
            }
        }
        if quote.is_some() || escaped {
            return None;
        }
        out.push(input[..end].to_string());
        input = &input[end..];
    }
    Some(out)
}

fn printf_missing_default_args(fmt: &str, provided: usize) -> Vec<Expression> {
    let defaults = printf_conversions(fmt)
        .into_iter()
        .map(printf_default_for_conversion)
        .collect::<Vec<_>>();
    if provided >= defaults.len() {
        Vec::new()
    } else {
        defaults.into_iter().skip(provided).collect()
    }
}

fn printf_conversions(fmt: &str) -> Vec<char> {
    let mut conversions = Vec::new();
    let chars: Vec<char> = fmt.chars().collect();
    let mut i = 0usize;
    while i < chars.len() {
        if chars[i] != '%' {
            i += 1;
            continue;
        }
        i += 1;
        if i >= chars.len() {
            break;
        }
        if chars[i] == '%' {
            i += 1;
            continue;
        }
        while i < chars.len() && matches!(chars[i], '#' | '0' | '-' | '+' | ' ' | '\'') {
            i += 1;
        }
        while i < chars.len() && chars[i].is_ascii_digit() {
            i += 1;
        }
        if i < chars.len() && chars[i] == '.' {
            i += 1;
            while i < chars.len() && chars[i].is_ascii_digit() {
                i += 1;
            }
        }
        while i < chars.len() && matches!(chars[i], 'h' | 'l' | 'L' | 'j' | 't' | 'z') {
            i += 1;
        }
        if i >= chars.len() {
            break;
        }
        let conv = chars[i];
        i += 1;
        conversions.push(conv);
    }
    conversions
}

fn printf_common_format(fmt: &str) -> String {
    let mut out = String::new();
    let chars: Vec<char> = fmt.chars().collect();
    let mut i = 0usize;
    while i < chars.len() {
        if chars[i] != '%' {
            out.push(chars[i]);
            i += 1;
            continue;
        }
        let start = i;
        i += 1;
        if i >= chars.len() {
            out.push('%');
            break;
        }
        if chars[i] == '%' {
            out.push('%');
            out.push('%');
            i += 1;
            continue;
        }
        while i < chars.len() && matches!(chars[i], '#' | '0' | '-' | '+' | ' ' | '\'') {
            i += 1;
        }
        while i < chars.len() && chars[i].is_ascii_digit() {
            i += 1;
        }
        if i < chars.len() && chars[i] == '.' {
            i += 1;
            while i < chars.len() && chars[i].is_ascii_digit() {
                i += 1;
            }
        }
        while i < chars.len() && matches!(chars[i], 'h' | 'l' | 'L' | 'j' | 't' | 'z') {
            i += 1;
        }
        if i >= chars.len() {
            out.extend(chars[start..].iter());
            break;
        }
        out.extend(chars[start..i].iter());
        out.push(if chars[i] == 'b' { 's' } else { chars[i] });
        i += 1;
    }
    out
}

fn printf_default_for_conversion(conv: char) -> Expression {
    match conv {
        'd' | 'i' | 'o' | 'u' | 'x' | 'X' | 'f' | 'F' | 'e' | 'E' | 'g' | 'G' | 'a' | 'A' => {
            int(0)
        }
        _ => lit(""),
    }
}

fn convert_bash_printf_arg(conv: char, value: Expression) -> Expression {
    if conv == 'b' {
        if let Some(text) = literal_string(&value) {
            return lit(&decode_printf_escapes(text));
        }
    }
    value
}

fn eval_const_arith(input: &str) -> Option<i64> {
    struct P<'a> {
        chars: Vec<char>,
        pos: usize,
        _src: &'a str,
    }
    impl<'a> P<'a> {
        fn new(src: &'a str) -> Self {
            Self {
                chars: src.chars().collect(),
                pos: 0,
                _src: src,
            }
        }
        fn skip_ws(&mut self) {
            while self.pos < self.chars.len() && self.chars[self.pos].is_whitespace() {
                self.pos += 1;
            }
        }
        fn eat(&mut self, ch: char) -> bool {
            self.skip_ws();
            if self.chars.get(self.pos) == Some(&ch) {
                self.pos += 1;
                true
            } else {
                false
            }
        }
        fn expr(&mut self) -> Option<i64> {
            let mut acc = self.term()?;
            loop {
                if self.eat('+') {
                    acc = acc.wrapping_add(self.term()?);
                } else if self.eat('-') {
                    acc = acc.wrapping_sub(self.term()?);
                } else {
                    return Some(acc);
                }
            }
        }
        fn term(&mut self) -> Option<i64> {
            let mut acc = self.factor()?;
            loop {
                if self.eat('*') {
                    acc = acc.wrapping_mul(self.factor()?);
                } else if self.eat('/') {
                    let rhs = self.factor()?;
                    if rhs == 0 {
                        return None;
                    }
                    acc /= rhs;
                } else if self.eat('%') {
                    let rhs = self.factor()?;
                    if rhs == 0 {
                        return None;
                    }
                    acc %= rhs;
                } else {
                    return Some(acc);
                }
            }
        }
        fn factor(&mut self) -> Option<i64> {
            self.skip_ws();
            if self.eat('(') {
                let value = self.expr()?;
                self.eat(')').then_some(value)
            } else {
                let sign = if self.eat('-') { -1 } else { 1 };
                self.skip_ws();
                let start = self.pos;
                while self.pos < self.chars.len()
                    && (self.chars[self.pos].is_ascii_alphanumeric()
                        || matches!(self.chars[self.pos], '#' | '@' | '_'))
                {
                    self.pos += 1;
                }
                if start == self.pos {
                    return None;
                }
                let token = self.chars[start..self.pos].iter().collect::<String>();
                let token = if sign < 0 {
                    format!("-{token}")
                } else {
                    token
                };
                parse_bash_integer(&token)
            }
        }
    }
    let mut p = P::new(input);
    let value = p.expr()?;
    p.skip_ws();
    (p.pos == p.chars.len()).then_some(value)
}

/// The raw text of a glob word whose quoted parts are all literal: quoted
/// segments have their metacharacters escaped so they match literally.
fn glob_word_text(pair: &Pair<Rule>) -> Option<String> {
    let mut out = String::new();
    for p in pair.clone().into_inner() {
        match p.as_rule() {
            Rule::bare_word | Rule::extglob_word => out.push_str(p.as_str()),
            Rule::single_quoted_string => {
                let s = p.as_str();
                out.push_str(&glob_escape(&s[1..s.len() - 1]));
            }
            Rule::ansi_c_quoting => {
                let s = p.as_str();
                out.push_str(&glob_escape(&decode_ansi_c(&s[2..s.len() - 1])));
            }
            Rule::quoted_string => {
                for q in p.into_inner() {
                    match q.as_rule() {
                        Rule::dq_text => out.push_str(&glob_escape(q.as_str())),
                        Rule::dq_escape => out.push_str(&glob_escape(&q.as_str()[1..])),
                        _ => return None,
                    }
                }
            }
            _ => return None,
        }
    }
    Some(out)
}

/// Backslash-escape glob metacharacters so quoted text matches literally.
fn glob_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        if matches!(c, '*' | '?' | '[' | ']' | '\\') {
            out.push('\\');
        }
        out.push(c);
    }
    out
}

/// Escape regex metacharacters.
fn regex_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        if matches!(
            c,
            '.' | '*'
                | '+'
                | '?'
                | '^'
                | '$'
                | '('
                | ')'
                | '['
                | ']'
                | '{'
                | '}'
                | '|'
                | '\\'
                | '/'
        ) {
            out.push('\\');
        }
        out.push(c);
    }
    out
}

fn bash_replacement_text(s: &str, patsub_replacement: bool) -> String {
    let mut out = String::new();
    let mut chars = s.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '\\' => {
                if let Some(next) = chars.next() {
                    out.push(next);
                } else {
                    out.push('\\');
                }
            }
            '&' if patsub_replacement => out.push_str("$&"),
            '$' => out.push_str("$$"),
            _ => out.push(c),
        }
    }
    out
}

/// A bash glob pattern as a JS regular expression body (unanchored).
/// `shortest` makes `*` lazy, for `${var#pattern}`.
fn glob_to_regex(pat: &str, shortest: bool) -> Option<String> {
    let chars: Vec<char> = pat.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '?' | '*' | '+' | '@' | '!' if i + 1 < chars.len() && chars[i + 1] == '(' => {
                let end = find_extglob_end(&chars, i + 1)?;
                let inner = &pat[char_to_byte(&chars, i + 2)..char_to_byte(&chars, end)];
                let mut alts = Vec::new();
                for alt in split_extglob_alternatives(inner) {
                    alts.push(glob_to_regex(alt, shortest)?);
                }
                let body = format!("(?:{})", alts.join("|"));
                match c {
                    '?' => out.push_str(&format!("{body}?")),
                    '*' => out.push_str(&format!("{body}*")),
                    '+' => out.push_str(&format!("{body}+")),
                    '@' => out.push_str(&body),
                    '!' => out.push_str(&format!("(?!{body}$)[\\s\\S]*")),
                    _ => unreachable!(),
                }
                i = end + 1;
            }
            '*' => {
                out.push_str(if shortest { "[\\s\\S]*?" } else { "[\\s\\S]*" });
                i += 1;
            }
            '?' => {
                out.push_str("[\\s\\S]");
                i += 1;
            }
            '\\' => {
                i += 1;
                if i < chars.len() {
                    out.push_str(&regex_escape(&chars[i].to_string()));
                    i += 1;
                }
            }
            '[' => {
                // Bracket expression.
                let mut j = i + 1;
                let mut class = String::from("[");
                if j < chars.len() && (chars[j] == '!' || chars[j] == '^') {
                    class.push('^');
                    j += 1;
                }
                if j < chars.len() && chars[j] == ']' {
                    class.push_str("\\]");
                    j += 1;
                }
                let mut closed = false;
                while j < chars.len() {
                    let d = chars[j];
                    if d == ']' {
                        closed = true;
                        j += 1;
                        break;
                    }
                    if d == '[' && j + 1 < chars.len() && chars[j + 1] == ':' {
                        if let Some(end) = pat[char_to_byte(&chars, j)..].find(":]") {
                            let name =
                                &pat[char_to_byte(&chars, j) + 2..char_to_byte(&chars, j) + end];
                            let expansion = match name {
                                "alpha" => "a-zA-Z",
                                "digit" => "0-9",
                                "alnum" => "a-zA-Z0-9",
                                "upper" => "A-Z",
                                "lower" => "a-z",
                                "space" => "\\s",
                                "blank" => " \\t",
                                "punct" => "!-\\/:-@\\[-`{-~",
                                "xdigit" => "0-9A-Fa-f",
                                "cntrl" => "\\x00-\\x1f",
                                "print" => "\\x20-\\x7e",
                                "graph" => "\\x21-\\x7e",
                                "word" => "\\w",
                                _ => return None,
                            };
                            class.push_str(expansion);
                            j += name.chars().count() + 4;
                            continue;
                        }
                    }
                    if d == '\\' && j + 1 < chars.len() {
                        class.push('\\');
                        class.push(chars[j + 1]);
                        j += 2;
                        continue;
                    }
                    if matches!(d, '^' | '\\' | '[') {
                        class.push('\\');
                    }
                    class.push(d);
                    j += 1;
                }
                if !closed {
                    // An unterminated `[` is a literal bracket.
                    out.push_str("\\[");
                    i += 1;
                    continue;
                }
                class.push(']');
                out.push_str(&class);
                i = j;
            }
            _ => {
                out.push_str(&regex_escape(&c.to_string()));
                i += 1;
            }
        }
    }
    Some(out)
}

fn find_extglob_end(chars: &[char], open: usize) -> Option<usize> {
    let mut depth = 1usize;
    let mut i = open + 1;
    while i < chars.len() {
        match chars[i] {
            '\\' => i += 2,
            '(' => {
                depth += 1;
                i += 1;
            }
            ')' => {
                depth -= 1;
                if depth == 0 {
                    return Some(i);
                }
                i += 1;
            }
            _ => i += 1,
        }
    }
    None
}

fn split_extglob_alternatives(inner: &str) -> Vec<&str> {
    let chars: Vec<char> = inner.chars().collect();
    let mut out = Vec::new();
    let mut depth = 0usize;
    let mut start = 0usize;
    let mut i = 0usize;
    while i < chars.len() {
        match chars[i] {
            '\\' => i += 2,
            '(' => {
                depth += 1;
                i += 1;
            }
            ')' => {
                depth = depth.saturating_sub(1);
                i += 1;
            }
            '|' if depth == 0 => {
                out.push(&inner[start..char_to_byte(&chars, i)]);
                i += 1;
                start = char_to_byte(&chars, i);
            }
            _ => i += 1,
        }
    }
    out.push(&inner[start..]);
    out
}

fn char_to_byte(chars: &[char], idx: usize) -> usize {
    chars[..idx].iter().map(|c| c.len_utf8()).sum()
}

/// A word with no quotes or expansions at all.
fn word_is_plain(pair: &Pair<Rule>) -> bool {
    pair.clone()
        .into_inner()
        .all(|p| p.as_rule() == Rule::bare_word)
}

fn word_has_quoted_part(pair: &Pair<Rule>) -> bool {
    if raw_word_has_shell_quote(pair.as_str()) {
        return true;
    }
    pair.clone().into_inner().any(|p| {
        matches!(
            p.as_rule(),
            Rule::quoted_string
                | Rule::single_quoted_string
                | Rule::ansi_c_quoting
                | Rule::locale_quoting
        )
    })
}

fn raw_word_has_shell_quote(raw: &str) -> bool {
    let mut escaped = false;
    for ch in raw.chars() {
        if escaped {
            escaped = false;
            continue;
        }
        match ch {
            '\\' => escaped = true,
            '"' | '\'' => return true,
            _ => {}
        }
    }
    false
}

fn word_starts_with_backslash(pair: &Pair<Rule>) -> bool {
    pair.as_rule() == Rule::word && pair.as_str().trim_start().starts_with('\\')
}

fn raw_may_contain_brace_expansion(raw: &str) -> bool {
    raw.contains('{') && raw.contains('}') && (raw.contains(',') || raw.contains(".."))
}

fn trim_shell_fragment(s: &str) -> &str {
    s.trim_matches(|c: char| c.is_whitespace() || c == ';' || c == '&')
}

fn find_next_if_boundary(s: &str) -> Option<(&'static str, usize)> {
    let mut best: Option<(&'static str, usize)> = None;
    for kw in ["elif", "else", "fi"] {
        if let Some(pos) = find_control_keyword(s, kw) {
            if best.is_none_or(|(_, best_pos)| pos < best_pos) {
                best = Some((kw, pos));
            }
        }
    }
    best
}

fn find_control_keyword(s: &str, keyword: &str) -> Option<usize> {
    let mut i = 0usize;
    let mut nested_if = 0usize;
    while let Some((pos, word)) = next_shell_word(s, i) {
        if word == "if" {
            nested_if += 1;
        } else if word == "fi" && nested_if > 0 {
            nested_if -= 1;
        } else if nested_if == 0 && word == keyword {
            return Some(pos);
        }
        i = pos + word.len();
    }
    None
}

fn next_shell_word(s: &str, mut i: usize) -> Option<(usize, &str)> {
    let bytes = s.as_bytes();
    while i < bytes.len() {
        match bytes[i] {
            b'\'' => i = skip_single_quote(bytes, i)?,
            b'"' => i = skip_double_quote(bytes, i)?,
            b'`' => i = skip_backtick(bytes, i)?,
            b'$' if bytes.get(i + 1) == Some(&b'(') => i = skip_balanced_parens(bytes, i + 1)?,
            b'$' if bytes.get(i + 1) == Some(&b'{') => i = skip_braced_param_bytes(bytes, i + 2)?,
            b'<' | b'>' if bytes.get(i + 1) == Some(&b'(') => {
                i = skip_balanced_parens(bytes, i + 1)?
            }
            b'(' if bytes.get(i + 1) == Some(&b'(') => i = skip_double_close(bytes, i + 2, b')')?,
            b'[' if bytes.get(i + 1) == Some(&b'[') => i = skip_double_close(bytes, i + 2, b']')?,
            b if is_shell_word_start(b) && word_left_boundary(bytes, i) => {
                let start = i;
                i += 1;
                while i < bytes.len() && is_shell_word_continue(bytes[i]) {
                    i += 1;
                }
                return Some((start, &s[start..i]));
            }
            _ => i += 1,
        }
    }
    None
}

fn word_left_boundary(bytes: &[u8], i: usize) -> bool {
    i == 0 || !is_shell_word_continue(bytes[i - 1])
}

fn is_shell_word_start(b: u8) -> bool {
    b.is_ascii_alphabetic() || b == b'_'
}

fn is_shell_word_continue(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}

fn skip_single_quote(bytes: &[u8], mut i: usize) -> Option<usize> {
    i += 1;
    while i < bytes.len() {
        if bytes[i] == b'\'' {
            return Some(i + 1);
        }
        i += 1;
    }
    Some(bytes.len())
}

fn skip_double_quote(bytes: &[u8], mut i: usize) -> Option<usize> {
    i += 1;
    while i < bytes.len() {
        match bytes[i] {
            b'\\' => i += 2,
            b'"' => return Some(i + 1),
            b'$' if bytes.get(i + 1) == Some(&b'(') => i = skip_balanced_parens(bytes, i + 1)?,
            b'$' if bytes.get(i + 1) == Some(&b'{') => i = skip_braced_param_bytes(bytes, i + 2)?,
            b'`' => i = skip_backtick(bytes, i)?,
            _ => i += 1,
        }
    }
    Some(bytes.len())
}

fn skip_backtick(bytes: &[u8], mut i: usize) -> Option<usize> {
    i += 1;
    while i < bytes.len() {
        match bytes[i] {
            b'\\' => i += 2,
            b'`' => return Some(i + 1),
            _ => i += 1,
        }
    }
    Some(bytes.len())
}

fn skip_balanced_parens(bytes: &[u8], mut i: usize) -> Option<usize> {
    let mut depth = 0usize;
    while i < bytes.len() {
        match bytes[i] {
            b'\'' => i = skip_single_quote(bytes, i)?,
            b'"' => i = skip_double_quote(bytes, i)?,
            b'`' => i = skip_backtick(bytes, i)?,
            b'(' => {
                depth += 1;
                i += 1;
            }
            b')' => {
                depth = depth.checked_sub(1)?;
                i += 1;
                if depth == 0 {
                    return Some(i);
                }
            }
            b'\\' => i += 2,
            _ => i += 1,
        }
    }
    Some(bytes.len())
}

fn skip_double_close(bytes: &[u8], mut i: usize, close: u8) -> Option<usize> {
    while i + 1 < bytes.len() {
        match bytes[i] {
            b'\'' => i = skip_single_quote(bytes, i)?,
            b'"' => i = skip_double_quote(bytes, i)?,
            b'`' => i = skip_backtick(bytes, i)?,
            b'$' if bytes.get(i + 1) == Some(&b'(') => i = skip_balanced_parens(bytes, i + 1)?,
            b'\\' => i += 2,
            b if b == close && bytes.get(i + 1) == Some(&close) => return Some(i + 2),
            _ => i += 1,
        }
    }
    Some(bytes.len())
}

fn brace_expand_text(raw: &str) -> Option<Vec<String>> {
    let (open, close) = find_brace_expansion_span(raw)?;
    let prefix = &raw[..open];
    let body = &raw[open + 1..close];
    let suffix = &raw[close + 1..];
    let elements = brace_body_elements(body)?;
    let suffixes = brace_expand_text(suffix).unwrap_or_else(|| vec![suffix.to_string()]);
    let mut out = Vec::new();
    for element in elements {
        let expanded_elements = brace_expand_text(&element).unwrap_or_else(|| vec![element]);
        for expanded_element in expanded_elements {
            for suffix in &suffixes {
                out.push(format!("{prefix}{expanded_element}{suffix}"));
            }
        }
    }
    Some(out)
}

fn find_brace_expansion_span(raw: &str) -> Option<(usize, usize)> {
    let bytes = raw.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        match bytes[i] {
            b'\\' => i += 2,
            b'$' if bytes.get(i + 1) == Some(&b'{') => {
                i = skip_braced_param_bytes(bytes, i + 2)?;
            }
            b'{' => {
                if let Some(close) = matching_brace_end(bytes, i) {
                    let body = &raw[i + 1..close];
                    if brace_body_elements(body).is_some() {
                        return Some((i, close));
                    }
                }
                i += 1;
            }
            _ => i += 1,
        }
    }
    None
}

fn matching_brace_end(bytes: &[u8], open: usize) -> Option<usize> {
    let mut depth = 0usize;
    let mut i = open;
    while i < bytes.len() {
        match bytes[i] {
            b'\\' => i += 2,
            b'$' if bytes.get(i + 1) == Some(&b'{') => {
                i = skip_braced_param_bytes(bytes, i + 2)?;
            }
            b'{' => {
                depth += 1;
                i += 1;
            }
            b'}' => {
                depth = depth.checked_sub(1)?;
                if depth == 0 {
                    return Some(i);
                }
                i += 1;
            }
            _ => i += 1,
        }
    }
    None
}

fn skip_braced_param_bytes(bytes: &[u8], mut i: usize) -> Option<usize> {
    while i < bytes.len() {
        match bytes[i] {
            b'\\' => i += 2,
            b'}' => return Some(i + 1),
            _ => i += 1,
        }
    }
    None
}

fn brace_body_elements(body: &str) -> Option<Vec<String>> {
    if let Some((start, end, step)) = parse_brace_range(body) {
        return Some(brace_range_elements(&start, &end, step));
    }
    let mut elements = Vec::new();
    let mut start = 0usize;
    let mut depth = 0usize;
    let mut saw_comma = false;
    let bytes = body.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        match bytes[i] {
            b'\\' => i += 2,
            b'$' if bytes.get(i + 1) == Some(&b'{') => {
                i = skip_braced_param_bytes(bytes, i + 2)?;
            }
            b'{' => {
                depth += 1;
                i += 1;
            }
            b'}' if depth > 0 => {
                depth -= 1;
                i += 1;
            }
            b',' if depth == 0 => {
                saw_comma = true;
                elements.push(body[start..i].to_string());
                start = i + 1;
                i += 1;
            }
            _ => i += 1,
        }
    }
    if !saw_comma {
        return None;
    }
    elements.push(body[start..].to_string());
    Some(elements)
}

fn parse_brace_range(body: &str) -> Option<(String, String, Option<i64>)> {
    if body.contains(',') {
        return None;
    }
    let mut parts = body.split("..");
    let start = parts.next()?.to_string();
    let end = parts.next()?.to_string();
    let step = parts.next().and_then(|part| part.parse::<i64>().ok());
    if parts.next().is_some() || start.is_empty() || end.is_empty() {
        return None;
    }
    let numeric = start.parse::<i64>().is_ok() && end.parse::<i64>().is_ok();
    let chars = start.chars().count() == 1
        && end.chars().count() == 1
        && start.chars().all(|c| c.is_ascii_alphabetic())
        && end.chars().all(|c| c.is_ascii_alphabetic());
    if !numeric && !chars {
        return None;
    }
    Some((start, end, step))
}

fn brace_range_elements(start: &str, end: &str, step: Option<i64>) -> Vec<String> {
    if let (Ok(a), Ok(b)) = (start.parse::<i64>(), end.parse::<i64>()) {
        let width = if (start.starts_with('0') && start.len() > 1)
            || (end.starts_with('0') && end.len() > 1)
        {
            start.len().max(end.len())
        } else {
            0
        };
        let step = step.unwrap_or(1).abs().max(1);
        let mut out = Vec::new();
        let mut value = a;
        loop {
            if (a <= b && value > b) || (a > b && value < b) {
                break;
            }
            out.push(if width > 0 {
                format!("{value:0width$}")
            } else {
                value.to_string()
            });
            value += if a <= b { step } else { -step };
        }
        return out;
    }
    let Some(a) = start.chars().next() else {
        return Vec::new();
    };
    let Some(b) = end.chars().next() else {
        return Vec::new();
    };
    let step = step.unwrap_or(1).unsigned_abs().max(1) as u32;
    let (a, b) = (a as u32, b as u32);
    let mut out = Vec::new();
    let mut value = a;
    loop {
        if (a <= b && value > b) || (a > b && value < b) {
            break;
        }
        out.push(char::from_u32(value).unwrap_or('?').to_string());
        if a <= b {
            value += step;
        } else {
            value = value.wrapping_sub(step);
        }
    }
    out
}

fn word_has_runtime_side_effect(pair: &Pair<Rule>) -> bool {
    match pair.as_rule() {
        Rule::dollar_paren_subst
        | Rule::backtick_subst
        | Rule::brace_command_subst
        | Rule::process_substitution => true,
        Rule::arithmetic_expansion => {
            arithmetic_expansion_inner(pair.as_str()).is_some_and(|inner| {
                inner.contains("++")
                    || inner.contains("--")
                    || inner.contains("+=")
                    || inner.contains("-=")
                    || inner.contains("*=")
                    || inner.contains("/=")
                    || inner.contains("%=")
                    || inner.contains("<<=")
                    || inner.contains(">>=")
                    || inner.contains("&=")
                    || inner.contains("|=")
                    || inner.contains("^=")
                    || inner.contains('=')
            })
        }
        Rule::braced_param => {
            let raw = pair.as_str();
            raw.contains(":=") || raw.contains('=')
        }
        _ => pair
            .clone()
            .into_inner()
            .any(|inner| word_has_runtime_side_effect(&inner)),
    }
}

fn word_is_unquoted_expansion_only(pair: &Pair<Rule>) -> bool {
    if pair.as_rule() != Rule::word || word_has_quoted_part(pair) {
        return false;
    }
    let mut saw_expansion = false;
    for part in pair.clone().into_inner() {
        match part.as_rule() {
            Rule::simple_param
            | Rule::braced_param
            | Rule::dollar_paren_subst
            | Rule::backtick_subst
            | Rule::brace_command_subst
            | Rule::arithmetic_expansion => saw_expansion = true,
            _ => return false,
        }
    }
    saw_expansion
}

fn word_has_unquoted_expansion(pair: &Pair<Rule>) -> bool {
    if word_has_quoted_part(pair) {
        return false;
    }
    match pair.as_rule() {
        Rule::simple_param
        | Rule::braced_param
        | Rule::dollar_paren_subst
        | Rule::backtick_subst
        | Rule::brace_command_subst
        | Rule::arithmetic_expansion => true,
        _ => pair
            .clone()
            .into_inner()
            .any(|inner| word_has_unquoted_expansion(&inner)),
    }
}

fn word_may_expand_to_multiple_args(pair: &Pair<Rule>) -> bool {
    if pair.as_rule() != Rule::word || word_has_quoted_part(pair) {
        return false;
    }
    let raw = pair.as_str().trim();
    if matches!(raw, "$@" | "$*") || whole_unquoted_param_name(raw).is_some() {
        return true;
    }
    if let Some(inner) = raw.strip_prefix("${").and_then(|s| s.strip_suffix('}')) {
        if inner
            .strip_suffix("[@]")
            .or_else(|| inner.strip_suffix("[*]"))
            .is_some_and(is_name)
        {
            return true;
        }
    }
    let mut parts = pair.clone().into_inner();
    if let Some(part) = parts.next() {
        if parts.next().is_none()
            && matches!(
                part.as_rule(),
                Rule::dollar_paren_subst | Rule::backtick_subst | Rule::brace_command_subst
            )
        {
            return true;
        }
    }
    glob_word_text(pair).is_some_and(|text| word_text_has_glob_meta(&text))
}

fn word_source_text(pair: &Pair<Rule>) -> String {
    pair.as_str().to_string()
}

fn whole_unquoted_param_name(raw: &str) -> Option<&str> {
    if raw.starts_with('"') || raw.starts_with('\'') {
        return None;
    }
    if let Some(name) = raw.strip_prefix('$') {
        return is_name(name).then_some(name);
    }
    let inner = raw.strip_prefix("${").and_then(|s| s.strip_suffix('}'))?;
    is_name(inner).then_some(inner)
}

fn whole_unquoted_special_param(raw: &str) -> Option<&str> {
    if raw.starts_with('"') || raw.starts_with('\'') {
        return None;
    }
    if let Some(name) = raw.strip_prefix('$') {
        return matches!(name, "?" | "!" | "$" | "-" | "_" | "#").then_some(name);
    }
    let inner = raw.strip_prefix("${").and_then(|s| s.strip_suffix('}'))?;
    matches!(inner, "?" | "!" | "$" | "-" | "_" | "#").then_some(inner)
}

fn source_has_positional_reference(source: &str) -> bool {
    let chars: Vec<char> = source.chars().collect();
    let mut i = 0usize;
    while i < chars.len() {
        if chars[i] != '$' {
            i += 1;
            continue;
        }
        let Some(next) = chars.get(i + 1).copied() else {
            break;
        };
        if next.is_ascii_digit() || matches!(next, '#' | '@' | '*') {
            return true;
        }
        if next == '{' {
            let mut j = i + 2;
            if matches!(chars.get(j), Some('!')) {
                j += 1;
            }
            if matches!(chars.get(j), Some(c) if c.is_ascii_digit() || matches!(c, '#' | '@' | '*'))
            {
                return true;
            }
        }
        i += 1;
    }
    false
}

fn function_source_preserves_static_state(source: &str) -> bool {
    if source.contains('=')
        || source.contains("++")
        || source.contains("--")
        || source.contains("<<")
        || source.contains(">>")
        || source.contains(">|")
    {
        return false;
    }
    for word in shell_words_for_detection(source) {
        if matches!(
            word.as_str(),
            "alias"
                | "cd"
                | "declare"
                | "eval"
                | "export"
                | "local"
                | "mapfile"
                | "read"
                | "readarray"
                | "readonly"
                | "set"
                | "shift"
                | "source"
                | "typeset"
                | "unalias"
                | "unset"
                | "."
        ) {
            return false;
        }
    }
    true
}

fn shell_words_for_detection(source: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut current = String::new();
    for ch in source.chars() {
        if ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' {
            current.push(ch);
        } else if !current.is_empty() {
            out.push(std::mem::take(&mut current));
        }
    }
    if !current.is_empty() {
        out.push(current);
    }
    out
}

fn static_dollar_name(chars: &[char], start: usize) -> Option<(String, usize)> {
    let next = *chars.get(start + 1)?;
    if matches!(next, '@' | '*') {
        return Some((next.to_string(), start + 2));
    }
    if next == '{' {
        let mut i = start + 2;
        let mut name = String::new();
        while i < chars.len() && chars[i] != '}' {
            name.push(chars[i]);
            i += 1;
        }
        if i >= chars.len() || !is_name(&name) {
            return None;
        }
        return Some((name, i + 1));
    }
    if !(next.is_ascii_alphabetic() || next == '_') {
        return None;
    }
    let mut i = start + 1;
    let mut name = String::new();
    while i < chars.len() && (chars[i].is_ascii_alphanumeric() || chars[i] == '_') {
        name.push(chars[i]);
        i += 1;
    }
    Some((name, i))
}

fn simple_braced_param_name(raw: &str) -> Option<&str> {
    let inner = raw.strip_prefix("${").and_then(|s| s.strip_suffix('}'))?;
    is_name(inner).then_some(inner)
}

fn word_text_has_glob_meta(s: &str) -> bool {
    let mut chars = s.chars();
    while let Some(c) = chars.next() {
        match c {
            '\\' => {
                chars.next();
            }
            '*' | '?' => return true,
            '[' => {
                let mut escaped = false;
                for ch in chars.by_ref() {
                    if escaped {
                        escaped = false;
                        continue;
                    }
                    if ch == '\\' {
                        escaped = true;
                        continue;
                    }
                    if ch == ']' {
                        return true;
                    }
                }
                return false;
            }
            _ => {}
        }
    }
    false
}

/// Does an unquoted part of the word contain a glob metacharacter?
fn word_has_unquoted_glob(pair: &Pair<Rule>) -> bool {
    pair.clone().into_inner().any(|p| match p.as_rule() {
        Rule::bare_word => {
            let s = p.as_str();
            word_text_has_glob_meta(s)
        }
        Rule::extglob_word => true,
        _ => false,
    })
}

fn unescape_bare(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.next() {
                Some('\n') => {}
                Some(n) => out.push(n),
                None => out.push('\\'),
            }
        } else {
            out.push(c);
        }
    }
    out
}

fn unescape_backtick(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.peek() {
                Some('`') | Some('$') | Some('\\') => {
                    out.push(chars.next().unwrap());
                }
                _ => out.push('\\'),
            }
        } else {
            out.push(c);
        }
    }
    out
}

/// `$'…'` escapes.
fn decode_ansi_c(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c != '\\' || i + 1 >= chars.len() {
            out.push(c);
            i += 1;
            continue;
        }
        let e = chars[i + 1];
        i += 2;
        match e {
            'n' => out.push('\n'),
            't' => out.push('\t'),
            'r' => out.push('\r'),
            'a' => out.push('\u{7}'),
            'b' => out.push('\u{8}'),
            'e' | 'E' => out.push('\u{1b}'),
            'f' => out.push('\u{c}'),
            'v' => out.push('\u{b}'),
            '\\' => out.push('\\'),
            '\'' => out.push('\''),
            '"' => out.push('"'),
            '?' => out.push('?'),
            'c' => {
                if i < chars.len() {
                    let ctl = chars[i].to_ascii_uppercase();
                    out.push(((ctl as u8) ^ 0x40) as char);
                    i += 1;
                }
            }
            'x' => {
                let mut v = 0u32;
                let mut n = 0;
                while n < 2 && i < chars.len() && chars[i].is_ascii_hexdigit() {
                    v = v * 16 + chars[i].to_digit(16).unwrap();
                    i += 1;
                    n += 1;
                }
                if v == 0 {
                    break;
                }
                out.push(char::from_u32(v).unwrap_or('?'));
            }
            'u' | 'U' => {
                let max = if e == 'u' { 4 } else { 8 };
                let mut v = 0u32;
                let mut n = 0;
                while n < max && i < chars.len() && chars[i].is_ascii_hexdigit() {
                    v = v * 16 + chars[i].to_digit(16).unwrap();
                    i += 1;
                    n += 1;
                }
                out.push(char::from_u32(v).unwrap_or('?'));
            }
            '0'..='7' => {
                let mut v = e.to_digit(8).unwrap();
                let mut n = 1;
                while n < 3 && i < chars.len() && chars[i].is_digit(8) {
                    v = v * 8 + chars[i].to_digit(8).unwrap();
                    i += 1;
                    n += 1;
                }
                if v == 0 {
                    break;
                }
                out.push(char::from_u32(v).unwrap_or('?'));
            }
            other => {
                out.push('\\');
                out.push(other);
            }
        }
    }
    out
}

/// Escapes in a `printf` FORMAT: like `$'…'`, but `%%` and other `%`
/// conversions are left to the formatter.
fn decode_printf_escapes(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c != '\\' || i + 1 >= chars.len() {
            out.push(c);
            i += 1;
            continue;
        }
        let e = chars[i + 1];
        i += 2;
        match e {
            'n' => out.push('\n'),
            't' => out.push('\t'),
            'r' => out.push('\r'),
            'a' => out.push('\u{7}'),
            'b' => out.push('\u{8}'),
            'e' | 'E' => out.push('\u{1b}'),
            'f' => out.push('\u{c}'),
            'v' => out.push('\u{b}'),
            '\\' => out.push('\\'),
            '\'' => out.push('\''),
            '"' => out.push('"'),
            'x' => {
                let mut v = 0u32;
                let mut n = 0;
                while n < 2 && i < chars.len() && chars[i].is_ascii_hexdigit() {
                    v = v * 16 + chars[i].to_digit(16).unwrap();
                    i += 1;
                    n += 1;
                }
                out.push(char::from_u32(v).unwrap_or('?'));
            }
            '0'..='7' => {
                let mut v = e.to_digit(8).unwrap();
                let mut n = 1;
                while n < 3 && i < chars.len() && chars[i].is_digit(8) {
                    v = v * 8 + chars[i].to_digit(8).unwrap();
                    i += 1;
                    n += 1;
                }
                out.push(char::from_u32(v).unwrap_or('?'));
            }
            other => {
                out.push('\\');
                out.push(other);
            }
        }
    }
    out
}

#[derive(Clone, Copy)]
enum BashIntegerError {
    InvalidBase,
    DigitOutOfRange,
    InvalidDigit,
}

/// bash integer literals: decimal, `0x` hex, leading-zero octal, `base#digits`.
fn parse_bash_integer(s: &str) -> Option<i64> {
    parse_bash_integer_result(s).ok()
}

fn parse_bash_integer_result(s: &str) -> Result<i64, BashIntegerError> {
    let t = s.trim();
    let (neg, t) = match t.strip_prefix('-') {
        Some(r) => (true, r),
        None => (false, t.strip_prefix('+').unwrap_or(t)),
    };
    let v = if let Some((base, digits)) = t.split_once('#') {
        let base: u32 = base.parse().map_err(|_| BashIntegerError::InvalidBase)?;
        if !(2..=64).contains(&base) || digits.is_empty() {
            return Err(BashIntegerError::InvalidBase);
        }
        let mut v: i64 = 0;
        for c in digits.chars() {
            let d = match c {
                '0'..='9' => c as u32 - '0' as u32,
                'a'..='z' => c as u32 - 'a' as u32 + 10,
                'A'..='Z' => {
                    if base <= 36 {
                        c as u32 - 'A' as u32 + 10
                    } else {
                        c as u32 - 'A' as u32 + 36
                    }
                }
                '@' => 62,
                '_' => 63,
                _ => return Err(BashIntegerError::InvalidDigit),
            };
            if d >= base {
                return Err(BashIntegerError::DigitOutOfRange);
            }
            v = v.wrapping_mul(base as i64).wrapping_add(d as i64);
        }
        v
    } else if let Some(h) = t.strip_prefix("0x").or_else(|| t.strip_prefix("0X")) {
        i64::from_str_radix(h, 16).map_err(|_| BashIntegerError::InvalidDigit)?
    } else if t.len() > 1 && t.starts_with('0') && t.chars().all(|c| c.is_ascii_digit()) {
        i64::from_str_radix(t, 8).map_err(|_| BashIntegerError::DigitOutOfRange)?
    } else if !t.is_empty() && t.chars().all(|c| c.is_ascii_digit()) {
        t.parse::<i64>()
            .map_err(|_| BashIntegerError::InvalidDigit)?
    } else {
        return Err(BashIntegerError::InvalidDigit);
    };
    Ok(if neg { -v } else { v })
}

fn bash_integer_error(s: &str) -> Option<BashIntegerError> {
    parse_bash_integer_result(s).err()
}

fn string_equals_any_expr(value: Expression, names: &[&str]) -> Expression {
    names
        .iter()
        .map(|name| binary(BinOp::StrictEq, value.clone(), lit(name)))
        .reduce(|left, right| binary(BinOp::Or, left, right))
        .unwrap_or_else(|| Expression::bool(false))
}

fn string_equals_any_owned_expr(value: Expression, names: &[String]) -> Expression {
    names
        .iter()
        .map(|name| binary(BinOp::StrictEq, value.clone(), lit(name)))
        .reduce(|left, right| binary(BinOp::Or, left, right))
        .unwrap_or_else(|| Expression::bool(false))
}

fn arithmetic_tokens(text: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut start: Option<usize> = None;
    let mut prev: Option<char> = None;
    for (idx, ch) in text.char_indices() {
        let is_token_char = ch.is_ascii_alphanumeric() || ch == '#' || ch == '@' || ch == '_';
        let sign_starts_number = matches!(ch, '+' | '-')
            && start.is_none()
            && text[idx + ch.len_utf8()..]
                .chars()
                .next()
                .is_some_and(|next| next.is_ascii_digit())
            && prev.is_none_or(|p| {
                p.is_whitespace()
                    || matches!(
                        p,
                        '(' | ','
                            | '?'
                            | ':'
                            | '+'
                            | '-'
                            | '*'
                            | '/'
                            | '%'
                            | '&'
                            | '|'
                            | '^'
                            | '<'
                            | '>'
                            | '='
                            | '!'
                    )
            });
        if is_token_char || sign_starts_number {
            if start.is_none() {
                start = Some(idx);
            }
        } else if let Some(s) = start.take() {
            out.push(&text[s..idx]);
        }
        prev = Some(ch);
    }
    if let Some(s) = start {
        out.push(&text[s..]);
    }
    out.into_iter()
        .filter(|token| {
            token
                .trim_start_matches(['+', '-'])
                .chars()
                .next()
                .is_some_and(|c| c.is_ascii_digit())
        })
        .collect()
}

/// A command word that can be a function name: mirror `function_name` in the
/// grammar, which is wider than shell variable identifiers.
fn is_function_name(name: &str) -> bool {
    let mut chars = name.chars();
    matches!(chars.next(), Some(c) if c.is_ascii_alphabetic() || c == '_')
        && chars.all(|c| c.is_ascii_alphanumeric() || matches!(c, '_' | '-' | '.' | ':'))
}

fn source_has_command_word(source: &str, needle: &str) -> bool {
    source
        .split(|c: char| !(c.is_ascii_alphanumeric() || matches!(c, '_' | '-' | '.' | ':')))
        .any(|word| word == needle)
}

fn has_keyword_function_definition(src: &str) -> bool {
    let src = src.trim_start();
    BASH_KEYWORD_NAMES
        .iter()
        .filter(|keyword| {
            keyword
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || c == '_')
        })
        .any(|keyword| {
            src.starts_with(&format!("{keyword}()"))
                || src
                    .strip_prefix("function ")
                    .is_some_and(|rest| rest.trim_start().starts_with(keyword))
        })
}

const BASH_KEYWORD_NAMES: &[&str] = &[
    "if", "then", "else", "elif", "fi", "for", "while", "until", "do", "done", "case", "esac",
    "function", "select", "time", "[[", "]]", "{", "}", "!", "coproc", "in",
];

const BASH_BUILTIN_NAMES: &[&str] = &[
    ":",
    ".",
    "[",
    "alias",
    "bg",
    "bind",
    "break",
    "builtin",
    "caller",
    "cd",
    "command",
    "compgen",
    "complete",
    "compopt",
    "continue",
    "declare",
    "dirs",
    "disown",
    "echo",
    "enable",
    "eval",
    "exec",
    "exit",
    "export",
    "false",
    "fc",
    "fg",
    "getopts",
    "hash",
    "help",
    "history",
    "jobs",
    "kill",
    "let",
    "local",
    "logout",
    "mapfile",
    "popd",
    "printf",
    "pushd",
    "pwd",
    "read",
    "readarray",
    "readonly",
    "return",
    "set",
    "shift",
    "shopt",
    "source",
    "suspend",
    "test",
    "times",
    "trap",
    "true",
    "type",
    "typeset",
    "ulimit",
    "umask",
    "unalias",
    "unset",
    "wait",
];
