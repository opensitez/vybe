//! C stdio.h — I/O normalisation adapters.
//!
//! printf/fprintf → puts(sprintf(...))  (sprintf itself built by the user's
//! emitter/sprintf.rs — not reimplemented here).
//! puts → wasi:cli:log via the "print" profile emit.

use crate::ast::{
    Argument, BinOp, BindingPattern, ExprKind, Expression, Literal, Modifiers, ObjectProperty,
    Param, PassBy, Statement, StmtKind, UnaryOp, VarDeclKind, VarDeclarator,
};

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn s(kind: StmtKind) -> Statement {
    Statement::new(kind)
}

fn ident(name: &str) -> Expression {
    e(ExprKind::Ident(name.to_string()))
}

fn member(object: Expression, field: &str) -> Expression {
    e(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn call_member(object: Expression, field: &str, args: Vec<Expression>) -> Expression {
    call(member(object, field), args)
}

fn bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    e(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn ternary(cond: Expression, then: Expression, else_: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then),
        else_: Box::new(else_),
    })
}

fn var_decl(name: &str, init: Expression) -> Statement {
    s(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: None,
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Var,
    })
}

fn if_stmt(cond: Expression, then_body: Vec<Statement>, else_body: Option<Vec<Statement>>) -> Statement {
    s(StmtKind::If { cond, then_body, elifs: Vec::new(), else_body })
}

fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    s(StmtKind::While { cond, body, else_body: None })
}

fn ret(value: Expression) -> Statement {
    s(StmtKind::Return(Some(value)))
}

fn expr_stmt(value: Expression) -> Statement {
    s(StmtKind::Expr(value))
}

fn function(name: &str, params: Vec<&str>, body: Vec<Statement>) -> Statement {
    s(StmtKind::FunctionDecl {
        name: name.to_string(),
        params: params
            .into_iter()
            .map(|param| Param {
                name: param.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect(),
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    })
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    e(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn lit_int(n: i64) -> Expression {
    e(ExprKind::Lit(Literal::Int(n)))
}

fn lit_float(n: f64) -> Expression {
    e(ExprKind::Lit(Literal::Float(n)))
}

fn lit_str(s: &str) -> Expression {
    e(ExprKind::Lit(Literal::Str(s.to_string())))
}

fn assign_expr(target: Expression, value: Expression) -> Expression {
    e(ExprKind::Assign {
        target: Box::new(target),
        value: Box::new(value),
    })
}

/// `printf(fmt, args...)` → `puts(sprintf(fmt, args...))`.
/// The caller strips the stream argument first for `fprintf`.
pub fn printf_to_puts(fmt: Expression, rest: Vec<Expression>) -> Expression {
    let mut sprintf_args = vec![fmt];
    sprintf_args.extend(rest);
    let sprintf_call = call(ident("sprintf"), sprintf_args);
    call(ident("puts"), vec![sprintf_call])
}

/// `fprintf(stream, fmt, args...)` → `puts(sprintf(fmt, args...))`.
/// Stream is dropped; output goes to the WASI log.
pub fn fprintf_to_puts(fmt: Expression, rest: Vec<Expression>) -> Expression {
    printf_to_puts(fmt, rest)
}

/// `sprintf(buf, fmt, args...)` → `buf = sprintf(fmt, args...)`.
/// The buffer target is returned as the assign target; the RHS is the call.
pub fn sprintf_assign(buf: Expression, fmt: Expression, rest: Vec<Expression>) -> Expression {
    let mut sprintf_args = vec![fmt];
    sprintf_args.extend(rest);
    let rhs = call(ident("sprintf"), sprintf_args);
    e(ExprKind::Assign {
        target: Box::new(buf),
        value: Box::new(rhs),
    })
}

/// C walker-compatible lowering: `printf(fmt, ...)` -> `__c_fputs_h(sprintf(...), 1)`.
pub fn printf_to_c_fputs(fmt: Expression, rest: Vec<Expression>) -> Expression {
    let mut sprintf_args = vec![fmt];
    sprintf_args.extend(rest);
    let rendered = call(ident("sprintf"), sprintf_args);
    call(ident("__c_fputs_h"), vec![rendered, lit_int(1)])
}

/// C walker-compatible lowering: `fprintf(file, fmt, ...)` -> `__c_fputs_h(sprintf(...), file)`.
pub fn fprintf_to_c_fputs(file: Expression, fmt: Expression, rest: Vec<Expression>) -> Expression {
    let mut sprintf_args = vec![fmt];
    sprintf_args.extend(rest);
    let rendered = call(ident("sprintf"), sprintf_args);
    call(ident("__c_fputs_h"), vec![rendered, file])
}

/// C walker-compatible lowering: `puts(text)` -> `__c_fputs_h(text + "\n", 1)`.
pub fn puts_to_c_fputs(text: Expression) -> Expression {
    call(
        ident("__c_fputs_h"),
        vec![
            e(ExprKind::Binary {
                op: crate::ast::BinOp::Add,
                left: Box::new(text),
                right: Box::new(lit_str("\n")),
            }),
            lit_int(1),
        ],
    )
}

/// Literal-only `sscanf` lowering used by C walker:
/// `sscanf("10 2.5", "%d %f", &a, &b)` -> sequence of assignments + count.
pub fn sscanf_literal(
    source_text: &str,
    format_text: &str,
    dest_targets: Vec<Expression>,
) -> Expression {
    let source_chars: Vec<char> = source_text.chars().collect();
    let format_chars: Vec<char> = format_text.chars().collect();
    let mut source_index = 0usize;
    let mut format_index = 0usize;
    let mut dest_index = 0usize;
    let mut count = 0i64;
    let mut stmts = Vec::new();

    let skip_source_ws = |source_index: &mut usize| {
        while *source_index < source_chars.len() && source_chars[*source_index].is_whitespace() {
            *source_index += 1;
        }
    };

    let parse_int_token = |token: &str, spec: char| -> Option<i64> {
        match spec {
            'd' => token.trim().parse::<i64>().ok(),
            'u' => token.trim().parse::<u64>().ok().map(|value| value as i64),
            'i' => {
                let trimmed = token.trim();
                let (sign, rest) = if let Some(rest) = trimmed.strip_prefix('-') {
                    (-1i64, rest)
                } else if let Some(rest) = trimmed.strip_prefix('+') {
                    (1i64, rest)
                } else {
                    (1i64, trimmed)
                };
                let value = if let Some(hex) = rest.strip_prefix("0x").or_else(|| rest.strip_prefix("0X")) {
                    i64::from_str_radix(hex, 16).ok()?
                } else if rest.starts_with('0') && rest.len() > 1 {
                    i64::from_str_radix(&rest[1..], 8).ok()?
                } else {
                    rest.parse::<i64>().ok()?
                };
                Some(sign * value)
            }
            _ => token.trim().parse::<i64>().ok(),
        }
    };

    while format_index < format_chars.len() && dest_index < dest_targets.len() {
        let fmt_ch = format_chars[format_index];
        format_index += 1;
        if fmt_ch.is_whitespace() {
            skip_source_ws(&mut source_index);
            continue;
        }
        if fmt_ch != '%' {
            if source_index < source_chars.len() && source_chars[source_index] == fmt_ch {
                source_index += 1;
                continue;
            }
            break;
        }

        let mut width_digits = String::new();
        while format_index < format_chars.len() && format_chars[format_index].is_ascii_digit() {
            width_digits.push(format_chars[format_index]);
            format_index += 1;
        }
        let width = if width_digits.is_empty() {
            None
        } else {
            width_digits.parse::<usize>().ok()
        };

        while format_index < format_chars.len()
            && matches!(format_chars[format_index], 'l' | 'h' | 'L')
        {
            format_index += 1;
        }

        if format_index >= format_chars.len() {
            break;
        }
        let spec = format_chars[format_index];
        format_index += 1;

        let target = dest_targets[dest_index].clone();
        dest_index += 1;

        match spec {
            'd' | 'u' | 'i' => {
                skip_source_ws(&mut source_index);
                let start = source_index;
                if source_index < source_chars.len()
                    && (source_chars[source_index] == '+' || source_chars[source_index] == '-')
                {
                    source_index += 1;
                }
                match spec {
                    'i' => {
                        if source_index + 1 < source_chars.len()
                            && source_chars[source_index] == '0'
                            && matches!(source_chars[source_index + 1], 'x' | 'X')
                        {
                            source_index += 2;
                            while source_index < source_chars.len()
                                && source_chars[source_index].is_ascii_hexdigit()
                            {
                                source_index += 1;
                            }
                        } else if source_index < source_chars.len() && source_chars[source_index] == '0' {
                            while source_index < source_chars.len()
                                && matches!(source_chars[source_index], '0'..='7')
                            {
                                source_index += 1;
                            }
                        } else {
                            while source_index < source_chars.len()
                                && source_chars[source_index].is_ascii_digit()
                            {
                                source_index += 1;
                            }
                        }
                    }
                    _ => {
                        while source_index < source_chars.len()
                            && source_chars[source_index].is_ascii_digit()
                        {
                            source_index += 1;
                        }
                    }
                }
                if source_index == start {
                    break;
                }
                let token: String = source_chars[start..source_index].iter().collect();
                if let Some(value) = parse_int_token(&token, spec) {
                    stmts.push(assign_expr(target, lit_int(value)));
                    count += 1;
                } else {
                    break;
                }
            }
            'f' => {
                skip_source_ws(&mut source_index);
                let start = source_index;
                while source_index < source_chars.len()
                    && matches!(source_chars[source_index], '0'..='9' | '+' | '-' | '.' | 'e' | 'E')
                {
                    source_index += 1;
                }
                if source_index == start {
                    break;
                }
                let token: String = source_chars[start..source_index].iter().collect();
                if let Ok(value) = token.trim().parse::<f64>() {
                    stmts.push(assign_expr(target, lit_float(value)));
                    count += 1;
                } else {
                    break;
                }
            }
            'c' => {
                if source_index >= source_chars.len() {
                    break;
                }
                let ch = source_chars[source_index];
                source_index += 1;
                stmts.push(assign_expr(target, lit_int(ch as i64)));
                count += 1;
            }
            's' => {
                skip_source_ws(&mut source_index);
                let start = source_index;
                while source_index < source_chars.len()
                    && !source_chars[source_index].is_whitespace()
                    && width.map(|limit| source_index - start < limit).unwrap_or(true)
                {
                    source_index += 1;
                }
                let token: String = source_chars[start..source_index].iter().collect();
                stmts.push(assign_expr(target, lit_str(&token)));
                count += 1;
            }
            '[' => {
                let mut negate = false;
                if format_index < format_chars.len() && format_chars[format_index] == '^' {
                    negate = true;
                    format_index += 1;
                }
                let mut stop_chars = Vec::new();
                while format_index < format_chars.len() && format_chars[format_index] != ']' {
                    stop_chars.push(format_chars[format_index]);
                    format_index += 1;
                }
                if format_index < format_chars.len() && format_chars[format_index] == ']' {
                    format_index += 1;
                }
                let start = source_index;
                while source_index < source_chars.len() {
                    let ch = source_chars[source_index];
                    let in_set = stop_chars.contains(&ch);
                    let should_stop = if negate { in_set } else { !in_set };
                    if should_stop || width.map(|limit| source_index - start >= limit).unwrap_or(false) {
                        break;
                    }
                    source_index += 1;
                }
                let token: String = source_chars[start..source_index].iter().collect();
                stmts.push(assign_expr(target, lit_str(&token)));
                count += 1;
            }
            _ => {}
        }
    }

    stmts.push(lit_int(count));
    if stmts.len() == 1 {
        stmts.pop().unwrap()
    } else {
        e(ExprKind::Sequence(stmts))
    }
}

// ── stdin token reader (libc surface, WASI-backed) ──────────────────────────
// Any libc-targeting language gets real stdin by including these runtime helpers
// in its prelude. The line-read primitive is the `__libc_readline` builtin, which
// each profile binds to `intrinsic:readline` → `emitter::io::emit_input` →
// `wasi:cli/stdin.get-stdin` + `wasi:io/streams.blocking-read`. So the C-visible
// `scanf`/`getchar`/`fgets` surface is satisfied entirely by WASM/WASI under the
// hood — no language-specific input host fn.

fn stdin_state() -> Expression {
    ident("__libc_stdin_state")
}
fn stdin_buf() -> Expression {
    member(stdin_state(), "buf")
}
fn stdin_eof() -> Expression {
    member(stdin_state(), "eof")
}

/// Runtime prelude statements implementing the WASI-backed stdin token reader.
/// Prepend these to any libc-targeting module's body.
pub fn stdin_runtime_helpers() -> Vec<Statement> {
    let mut out = Vec::new();

    // shared state: { buf: "", eof: 0, allow_blocking: 0 } — mutated via index assignment so the
    // helpers share one buffer (a bare scalar global reassigned inside a function
    // would shadow as a local).
    out.push(s(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("__libc_stdin_state".to_string()),
            type_hint: None,
            init: Some(e(ExprKind::Object(vec![
                ObjectProperty::KeyValue { key: lit_str("buf"), value: lit_str("") },
                ObjectProperty::KeyValue { key: lit_str("eof"), value: lit_int(0) },
                ObjectProperty::KeyValue { key: lit_str("allow_blocking"), value: lit_int(0) },
            ]))),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Var,
    }));

    // __libc_stdin_set_blocking(flag): opt-in to real blocking stdin reads.
    out.push(function(
        "__libc_stdin_set_blocking",
        vec!["flag"],
        vec![
            expr_stmt(assign_expr(member(stdin_state(), "allow_blocking"), ident("flag"))),
            ret(lit_int(0)),
        ],
    ));

    // __libc_stdin_read_line(): one line from WASI stdin via `intrinsic:readline`
    // (emitter::io::emit_input → wasi:cli/stdin.get-stdin +
    // wasi:io/streams.blocking-read).
    //
    // Default mode is non-blocking EOF for deterministic test/runtime behavior.
    // Call `__libc_stdin_set_blocking(1)` to opt into real blocking stdin reads.
    // Returns the line string, or "" when the read yields a non-string
    // (EOF / stream-error record) so the tokenizer sees a clean EOF.
    out.push(function(
        "__libc_stdin_read_line",
        vec![],
        vec![
            if_stmt(
                bin(BinOp::Eq, member(stdin_state(), "allow_blocking"), lit_int(0)),
                vec![ret(lit_str(""))],
                None,
            ),
            var_decl("__l", call(ident("__libc_readline"), vec![])),
            if_stmt(
                bin(
                    BinOp::Eq,
                    e(ExprKind::Unary { op: UnaryOp::Typeof, expr: Box::new(ident("__l")) }),
                    lit_str("string"),
                ),
                vec![ret(ident("__l"))],
                Some(vec![ret(lit_str(""))]),
            ),
        ],
    ));

    // __libc_stdin_token(): next whitespace-delimited token, or "" at EOF.
    out.push(function(
        "__libc_stdin_token",
        vec![],
        vec![
            var_decl("done", lit_int(0)),
            while_stmt(
                bin(BinOp::Eq, ident("done"), lit_int(0)),
                vec![
                    while_stmt(
                        bin(
                            BinOp::And,
                            bin(BinOp::Gt, member(stdin_buf(), "length"), lit_int(0)),
                            bin(BinOp::LtEq, call_member(stdin_buf(), "charCodeAt", vec![lit_int(0)]), lit_int(32)),
                        ),
                        vec![expr_stmt(assign_expr(stdin_buf(), call_member(stdin_buf(), "substring", vec![lit_int(1)])))],
                    ),
                    if_stmt(
                        bin(BinOp::Gt, member(stdin_buf(), "length"), lit_int(0)),
                        vec![expr_stmt(assign_expr(ident("done"), lit_int(1)))],
                        Some(vec![
                            if_stmt(
                                bin(BinOp::NotEq, stdin_eof(), lit_int(0)),
                                vec![ret(lit_str(""))],
                                None,
                            ),
                            var_decl("line", call(ident("__libc_stdin_read_line"), vec![])),
                            if_stmt(
                                bin(BinOp::Eq, member(ident("line"), "length"), lit_int(0)),
                                vec![expr_stmt(assign_expr(stdin_eof(), lit_int(1)))],
                                Some(vec![expr_stmt(assign_expr(
                                    stdin_buf(),
                                    bin(BinOp::Add, bin(BinOp::Add, stdin_buf(), ident("line")), lit_str(" ")),
                                ))]),
                            ),
                        ]),
                    ),
                ],
            ),
            var_decl("i", lit_int(0)),
            while_stmt(
                bin(
                    BinOp::And,
                    bin(BinOp::Lt, ident("i"), member(stdin_buf(), "length")),
                    bin(BinOp::Gt, call_member(stdin_buf(), "charCodeAt", vec![ident("i")]), lit_int(32)),
                ),
                vec![expr_stmt(assign_expr(ident("i"), bin(BinOp::Add, ident("i"), lit_int(1))))],
            ),
            var_decl("tok", call_member(stdin_buf(), "substring", vec![lit_int(0), ident("i")])),
            expr_stmt(assign_expr(stdin_buf(), call_member(stdin_buf(), "substring", vec![ident("i")]))),
            ret(ident("tok")),
        ],
    ));

    // __libc_stdin_char(): next single character, or "" at EOF.
    out.push(function(
        "__libc_stdin_char",
        vec![],
        vec![
            while_stmt(
                bin(
                    BinOp::And,
                    bin(BinOp::Eq, member(stdin_buf(), "length"), lit_int(0)),
                    bin(BinOp::Eq, stdin_eof(), lit_int(0)),
                ),
                vec![
                    var_decl("line", call(ident("__libc_stdin_read_line"), vec![])),
                    if_stmt(
                        bin(BinOp::Eq, member(ident("line"), "length"), lit_int(0)),
                        vec![expr_stmt(assign_expr(stdin_eof(), lit_int(1)))],
                        Some(vec![expr_stmt(assign_expr(stdin_buf(), bin(BinOp::Add, ident("line"), lit_str("\n"))))]),
                    ),
                ],
            ),
            if_stmt(
                bin(BinOp::Eq, member(stdin_buf(), "length"), lit_int(0)),
                vec![ret(lit_str(""))],
                None,
            ),
            var_decl("ch", call_member(stdin_buf(), "substring", vec![lit_int(0), lit_int(1)])),
            expr_stmt(assign_expr(stdin_buf(), call_member(stdin_buf(), "substring", vec![lit_int(1)]))),
            ret(ident("ch")),
        ],
    ));

    out
}

/// Lower `scanf(fmt, t1, ...)` (fmt a compile-time literal, targets already
/// address-stripped) into a sequence that reads conversions from the WASI-backed
/// stdin token reader, assigns each target, and evaluates to the match count.
/// Stops at the first failed conversion, per C semantics. `tmp_id` must be unique
/// per call site to avoid temp-name collisions.
pub fn scanf(fmt: &str, targets: Vec<Expression>, tmp_id: u32) -> Expression {
    let mut specs: Vec<char> = Vec::new();
    let chars: Vec<char> = fmt.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] != '%' {
            i += 1;
            continue;
        }
        i += 1;
        if i < chars.len() && chars[i] == '%' {
            i += 1;
            continue;
        }
        while i < chars.len() && chars[i].is_ascii_digit() {
            i += 1;
        }
        while i < chars.len() && matches!(chars[i], 'l' | 'h' | 'L' | 'z' | 'j' | 't') {
            i += 1;
        }
        if i < chars.len() {
            specs.push(chars[i]);
            i += 1;
        }
    }

    let n_var = format!("__sc_n{tmp_id}");
    let ok_var = format!("__sc_ok{tmp_id}");
    let tok_var = format!("__sc_tok{tmp_id}");

    let mut seq: Vec<Expression> = vec![
        assign_expr(ident(&n_var), lit_int(0)),
        assign_expr(ident(&ok_var), lit_int(1)),
    ];

    for (idx, spec) in specs.iter().enumerate() {
        let Some(target) = targets.get(idx).cloned() else {
            break;
        };
        let reader = if *spec == 'c' { "__libc_stdin_char" } else { "__libc_stdin_token" };
        seq.push(assign_expr(
            ident(&tok_var),
            ternary(ident(&ok_var), call(ident(reader), vec![]), lit_str("")),
        ));
        let converted = match spec {
            'd' | 'u' => bin(BinOp::Or, call(ident("parseInt"), vec![ident(&tok_var), lit_int(10)]), lit_int(0)),
            'i' => bin(BinOp::Or, call(ident("parseInt"), vec![ident(&tok_var)]), lit_int(0)),
            'x' | 'X' => bin(BinOp::Or, call(ident("parseInt"), vec![ident(&tok_var), lit_int(16)]), lit_int(0)),
            'o' => bin(BinOp::Or, call(ident("parseInt"), vec![ident(&tok_var), lit_int(8)]), lit_int(0)),
            'f' | 'e' | 'g' | 'F' | 'E' | 'G' | 'a' => {
                bin(BinOp::Or, call(ident("parseFloat"), vec![ident(&tok_var)]), lit_float(0.0))
            }
            _ => ident(&tok_var),
        };
        seq.push(ternary(
            bin(BinOp::Gt, member(ident(&tok_var), "length"), lit_int(0)),
            e(ExprKind::Sequence(vec![
                assign_expr(target, converted),
                assign_expr(ident(&n_var), bin(BinOp::Add, ident(&n_var), lit_int(1))),
            ])),
            assign_expr(ident(&ok_var), lit_int(0)),
        ));
    }

    seq.push(ident(&n_var));
    e(ExprKind::Sequence(seq))
}
