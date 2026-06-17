//! C stdio.h — I/O normalisation adapters.
//!
//! printf/fprintf → puts(sprintf(...))  (sprintf itself built by the user's
//! emitter/sprintf.rs — not reimplemented here).
//! puts → wasi:cli:log via the "print" profile emit.

use crate::ast::{Argument, ExprKind, Expression, Literal};

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn ident(name: &str) -> Expression {
    e(ExprKind::Ident(name.to_string()))
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
