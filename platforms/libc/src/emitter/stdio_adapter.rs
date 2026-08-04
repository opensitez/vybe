//! C stdio.h — I/O normalisation adapters.
//!
//! printf/fprintf → `__c_fputs_h(sprintf(...))`.
//! sprintf/sscanf formatting is owned by libc under `platforms/libc`.

use vybe_ast::{
    Argument, BinOp, BindingPattern, ExprKind, Expression, Literal, Modifiers, ObjectProperty,
    Param, PassBy, Statement, StmtKind, UnaryOp, VarDeclKind, VarDeclarator };

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
        null_safe: false })
}

fn call_member(object: Expression, field: &str, args: Vec<Expression>) -> Expression {
    call(member(object, field), args)
}

fn bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    e(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right) })
}

fn ternary(cond: Expression, then: Expression, else_: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then),
        else_: Box::new(else_) })
}

fn var_decl(name: &str, init: Expression) -> Statement {
    s(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: None,
            init: Some(init),
            array_bounds: None,
            with_events: false }],
        kind: VarDeclKind::Var })
}

fn if_stmt(
    cond: Expression,
    then_body: Vec<Statement>,
    else_body: Option<Vec<Statement>>,
) -> Statement {
    s(StmtKind::If {
        cond,
        then_body,
        elifs: Vec::new(),
        else_body })
}

fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    s(StmtKind::While {
        cond,
        body,
        else_body: None })
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
                is_nullable: false })
            .collect(),
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false })
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    e(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false })
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
        value: Box::new(value) })
}

const C_SPRINTF: &str = "__c_sprintf";

/// `printf(fmt, args...)` → `puts(__c_sprintf(fmt, args...))`.
/// The caller strips the stream argument first for `fprintf`.
pub fn printf_to_puts(fmt: Expression, rest: Vec<Expression>) -> Expression {
    let mut sprintf_args = vec![fmt];
    sprintf_args.extend(rest);
    let sprintf_call = call(ident(C_SPRINTF), sprintf_args);
    call(ident("puts"), vec![sprintf_call])
}

/// `fprintf(stream, fmt, args...)` → `puts(sprintf(fmt, args...))`.
/// Stream is dropped; output goes to the WASI log.
pub fn fprintf_to_puts(fmt: Expression, rest: Vec<Expression>) -> Expression {
    printf_to_puts(fmt, rest)
}

/// `sprintf(buf, fmt, args...)` → `buf = __c_sprintf(fmt, args...)`.
/// The buffer target is returned as the assign target; the RHS is the call.
pub fn sprintf_assign(buf: Expression, fmt: Expression, rest: Vec<Expression>) -> Expression {
    let mut sprintf_args = vec![fmt];
    sprintf_args.extend(rest);
    let rhs = call(ident(C_SPRINTF), sprintf_args);
    e(ExprKind::Assign {
        target: Box::new(buf),
        value: Box::new(rhs) })
}

/// C walker-compatible lowering: `printf(fmt, ...)` -> `__c_fputs_h(__c_sprintf(...), 1)`.
pub fn printf_to_c_fputs(fmt: Expression, rest: Vec<Expression>) -> Expression {
    let mut sprintf_args = vec![fmt];
    sprintf_args.extend(rest);
    let rendered = call(ident(C_SPRINTF), sprintf_args);
    call(ident("__c_fputs_h"), vec![rendered, lit_int(1)])
}

pub fn normalize_printf_literal_format(format_text: &str, arg_count: usize) -> String {
    let mut out = format_text.to_string();
    if arg_count == 0 {
        out = collapse_stray_triple_percents(&out);
    }
    out = normalize_integer_precision_padding(&out);
    out
}

fn normalize_integer_precision_padding(input: &str) -> String {
    let chars: Vec<char> = input.chars().collect();
    let mut out = String::new();
    let mut i = 0usize;
    while i < chars.len() {
        if chars[i] != '%' {
            out.push(chars[i]);
            i += 1;
            continue;
        }
        if i + 1 < chars.len() && chars[i + 1] == '%' {
            out.push('%');
            out.push('%');
            i += 2;
            continue;
        }
        let start = i;
        i += 1;
        while i < chars.len() && matches!(chars[i], '-' | '+' | ' ' | '#') {
            i += 1;
        }
        if i < chars.len() && chars[i] == '.' {
            let precision_start = i + 1;
            let mut j = precision_start;
            while j < chars.len() && chars[j].is_ascii_digit() {
                j += 1;
            }
            if j > precision_start && j < chars.len() && matches!(chars[j], 'd' | 'i' | 'u') {
                let precision: String = chars[precision_start..j].iter().collect();
                if precision != "0" {
                    out.extend(chars[start..precision_start - 1].iter());
                    out.push('0');
                    out.push_str(&precision);
                    out.push(chars[j]);
                    i = j + 1;
                    continue;
                }
            }
        }
        let mut end = i;
        while end < chars.len()
            && !matches!(
                chars[end],
                'd' | 'i'
                    | 'u'
                    | 'o'
                    | 'x'
                    | 'X'
                    | 'f'
                    | 'F'
                    | 'e'
                    | 'E'
                    | 'g'
                    | 'G'
                    | 'a'
                    | 'A'
                    | 'c'
                    | 's'
                    | 'p'
                    | 'n'
            )
        {
            end += 1;
        }
        if end < chars.len() {
            end += 1;
        }
        out.extend(chars[start..end].iter());
        i = end;
    }
    out
}

fn collapse_stray_triple_percents(input: &str) -> String {
    let mut out = String::new();
    let bytes = input.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] != b'%' {
            // Copy a whole UTF-8 CHARACTER. `bytes[index] as char` is a
            // Latin-1 decode: each byte of `é` became a separate char, so any
            // non-ASCII text in a `printf` FORMAT string came out as mojibake
            // (`printf("café\n")` printed `cafÃ©`) while `puts("café")` was
            // fine. unifiedstringplan.md step 0 — and a site OUTSIDE the 8
            // walkers that audit counted.
            let ch = input[index..]
                .chars()
                .next()
                .expect("index is a char boundary");
            out.push(ch);
            index += ch.len_utf8();
            continue;
        }
        let start = index;
        while index < bytes.len() && bytes[index] == b'%' {
            index += 1;
        }
        let count = index - start;
        if count == 3 {
            out.push_str("%%");
        } else {
            for _ in 0..count {
                out.push('%');
            }
        }
    }
    out
}

/// Literal-format `printf` lowering for `%n`.
///
/// The libc formatter returns a string, but `%n` is a C side effect: it stores
/// the number of characters emitted so far through the matching pointer
/// argument. For literal formats we split around `%n`, keep using libc
/// `sprintf` for each rendered segment, emit those segments with the same
/// `__c_fputs_h` path as printf, and assign the accumulated rendered length to
/// each `%n` target.
pub fn printf_with_n_to_c_fputs(
    format_text: &str,
    rest: Vec<Expression>,
    _file: Expression,
) -> Option<Expression> {
    let mut parser = PrintfNSplitter::new(format_text);
    let mut saw_n = false;
    let mut arg_index = 0usize;
    let mut segment_arg_start = 0usize;
    let mut segment = String::new();
    let mut seq = Vec::new();

    while let Some(item) = parser.next_item() {
        match item {
            PrintfItem::Literal(text) => segment.push_str(text),
            PrintfItem::Conversion { text, consumes_arg } => {
                segment.push_str(text);
                if consumes_arg {
                    arg_index += 1;
                }
            }
            PrintfItem::Count => {
                saw_n = true;
                flush_printf_segment(&mut seq, &mut segment, &rest, segment_arg_start, arg_index);
                if parser.remaining_starts_with_newline() {
                    segment.push('\n');
                    parser.skip_one_char();
                    flush_printf_segment(
                        &mut seq,
                        &mut segment,
                        &rest,
                        segment_arg_start,
                        arg_index,
                    );
                }
                if let Some(target) = rest.get(arg_index).cloned() {
                    seq.push(assign_expr(
                        pointer_write_target(target),
                        current_printf_n_count(),
                    ));
                }
                arg_index += 1;
                segment_arg_start = arg_index;
            }
        }
    }

    if !saw_n {
        return None;
    }

    flush_printf_segment(&mut seq, &mut segment, &rest, segment_arg_start, arg_index);

    if seq.is_empty() {
        Some(lit_int(0))
    } else if seq.len() == 1 {
        seq.pop()
    } else {
        Some(e(ExprKind::Sequence(seq)))
    }
}

/// C walker-compatible lowering: `fprintf(file, fmt, ...)` -> `__c_fputs_h(__c_sprintf(...), file)`.
pub fn fprintf_to_c_fputs(file: Expression, fmt: Expression, rest: Vec<Expression>) -> Expression {
    let mut sprintf_args = vec![fmt];
    sprintf_args.extend(rest);
    let rendered = call(ident(C_SPRINTF), sprintf_args);
    call(ident("__c_fputs_h"), vec![rendered, file])
}

fn flush_printf_segment(
    seq: &mut Vec<Expression>,
    segment: &mut String,
    rest: &[Expression],
    start: usize,
    end: usize,
) {
    if segment.is_empty() {
        return;
    }
    let rendered = sprintf_expr(lit_str(segment), rest, start, end);
    let rendered_len = member(rendered.clone(), "length");
    seq.push(call(ident("__c_fputs_h"), vec![rendered, lit_int(1)]));
    seq.push(assign_expr(
        ident("__c_printf_n_count"),
        bin(BinOp::Add, current_printf_n_count(), rendered_len),
    ));
    segment.clear();
}

fn current_printf_n_count() -> Expression {
    ternary(
        ident("__c_printf_n_count"),
        ident("__c_printf_n_count"),
        lit_int(0),
    )
}

fn sprintf_expr(fmt: Expression, rest: &[Expression], start: usize, end: usize) -> Expression {
    let mut args = vec![fmt];
    for arg in rest.iter().skip(start).take(end.saturating_sub(start)) {
        args.push(arg.clone());
    }
    call(ident(C_SPRINTF), args)
}

fn pointer_write_target(value: Expression) -> Expression {
    match value.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr } => *expr,
        other => e(other) }
}

enum PrintfItem<'a> {
    Literal(&'a str),
    Conversion { text: &'a str, consumes_arg: bool },
    Count }

struct PrintfNSplitter<'a> {
    text: &'a str,
    index: usize }

impl<'a> PrintfNSplitter<'a> {
    fn new(text: &'a str) -> Self {
        Self { text, index: 0 }
    }

    fn next_item(&mut self) -> Option<PrintfItem<'a>> {
        let bytes = self.text.as_bytes();
        if self.index >= bytes.len() {
            return None;
        }
        let start = self.index;
        if bytes[start] != b'%' {
            while self.index < bytes.len() && bytes[self.index] != b'%' {
                self.index += 1;
            }
            return Some(PrintfItem::Literal(&self.text[start..self.index]));
        }

        self.index += 1;
        if self.index >= bytes.len() {
            return Some(PrintfItem::Conversion {
                text: &self.text[start..self.index],
                consumes_arg: false });
        }
        if bytes[self.index] == b'%' {
            self.index += 1;
            return Some(PrintfItem::Conversion {
                text: &self.text[start..self.index],
                consumes_arg: false });
        }

        while self.index < bytes.len()
            && matches!(bytes[self.index], b'-' | b'+' | b' ' | b'#' | b'0')
        {
            self.index += 1;
        }
        if self.index < bytes.len() && bytes[self.index] == b'\'' {
            self.index += 1;
            if self.index < bytes.len() {
                self.index += 1;
            }
        }
        while self.index < bytes.len() && bytes[self.index].is_ascii_digit() {
            self.index += 1;
        }
        if self.index < bytes.len() && bytes[self.index] == b'.' {
            self.index += 1;
            while self.index < bytes.len() && bytes[self.index].is_ascii_digit() {
                self.index += 1;
            }
        }
        while self.index < bytes.len()
            && matches!(bytes[self.index], b'h' | b'l' | b'j' | b'z' | b't' | b'L')
        {
            self.index += 1;
        }
        if self.index >= bytes.len() {
            return Some(PrintfItem::Conversion {
                text: &self.text[start..self.index],
                consumes_arg: false });
        }

        let conv = bytes[self.index];
        self.index += 1;
        if conv == b'n' {
            Some(PrintfItem::Count)
        } else {
            Some(PrintfItem::Conversion {
                text: &self.text[start..self.index],
                consumes_arg: conv != b'%' })
        }
    }

    fn remaining_starts_with_newline(&self) -> bool {
        self.text[self.index..].starts_with('\n')
    }

    fn skip_one_char(&mut self) {
        if let Some(ch) = self.text[self.index..].chars().next() {
            self.index += ch.len_utf8();
        }
    }
}

/// C walker-compatible lowering: `puts(text)` -> `__c_fputs_h(text + "\n", 1)`.
pub fn puts_to_c_fputs(text: Expression) -> Expression {
    call(
        ident("__c_fputs_h"),
        vec![
            e(ExprKind::Binary {
                op: vybe_ast::BinOp::Add,
                left: Box::new(text),
                right: Box::new(lit_str("\n")) }),
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
    let mut input_failure = false;
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
            'o' => u64::from_str_radix(token.trim(), 8)
                .ok()
                .map(|value| value as i64),
            'x' | 'X' => {
                let trimmed = token.trim();
                let rest = trimmed
                    .strip_prefix("0x")
                    .or_else(|| trimmed.strip_prefix("0X"))
                    .unwrap_or(trimmed);
                u64::from_str_radix(rest, 16).ok().map(|value| value as i64)
            }
            'i' => {
                let trimmed = token.trim();
                let (sign, rest) = if let Some(rest) = trimmed.strip_prefix('-') {
                    (-1i64, rest)
                } else if let Some(rest) = trimmed.strip_prefix('+') {
                    (1i64, rest)
                } else {
                    (1i64, trimmed)
                };
                let value = if let Some(hex) =
                    rest.strip_prefix("0x").or_else(|| rest.strip_prefix("0X"))
                {
                    i64::from_str_radix(hex, 16).ok()?
                } else if rest.starts_with('0') && rest.len() > 1 {
                    i64::from_str_radix(&rest[1..], 8).ok()?
                } else {
                    rest.parse::<i64>().ok()?
                };
                Some(sign * value)
            }
            _ => token.trim().parse::<i64>().ok() }
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

        let suppress_assignment =
            if format_index < format_chars.len() && format_chars[format_index] == '*' {
                format_index += 1;
                true
            } else {
                false
            };

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
            && matches!(
                format_chars[format_index],
                'l' | 'h' | 'L' | 'z' | 'j' | 't'
            )
        {
            format_index += 1;
        }

        if format_index >= format_chars.len() {
            break;
        }
        let spec = format_chars[format_index];
        format_index += 1;

        let target = if suppress_assignment {
            lit_int(0)
        } else {
            let target = dest_targets[dest_index].clone();
            dest_index += 1;
            target
        };

        match spec {
            'd' | 'u' | 'i' | 'o' | 'x' | 'X' => {
                skip_source_ws(&mut source_index);
                if source_index >= source_chars.len() {
                    input_failure = true;
                    break;
                }
                let start = source_index;
                let limit = width
                    .map(|width| start.saturating_add(width).min(source_chars.len()))
                    .unwrap_or(source_chars.len());
                if source_index < source_chars.len()
                    && source_index < limit
                    && (source_chars[source_index] == '+' || source_chars[source_index] == '-')
                {
                    source_index += 1;
                }
                match spec {
                    'i' => {
                        if source_index + 1 < limit
                            && source_chars[source_index] == '0'
                            && matches!(source_chars[source_index + 1], 'x' | 'X')
                        {
                            source_index += 2;
                            while source_index < limit
                                && source_chars[source_index].is_ascii_hexdigit()
                            {
                                source_index += 1;
                            }
                        } else if source_index < limit && source_chars[source_index] == '0' {
                            while source_index < limit
                                && matches!(source_chars[source_index], '0'..='7')
                            {
                                source_index += 1;
                            }
                        } else {
                            while source_index < limit
                                && source_chars[source_index].is_ascii_digit()
                            {
                                source_index += 1;
                            }
                        }
                    }
                    'o' => {
                        while source_index < limit
                            && matches!(source_chars[source_index], '0'..='7')
                        {
                            source_index += 1;
                        }
                    }
                    'x' | 'X' => {
                        if source_index + 1 < limit
                            && source_chars[source_index] == '0'
                            && matches!(source_chars[source_index + 1], 'x' | 'X')
                        {
                            source_index += 2;
                        }
                        while source_index < limit && source_chars[source_index].is_ascii_hexdigit()
                        {
                            source_index += 1;
                        }
                    }
                    _ => {
                        while source_index < limit && source_chars[source_index].is_ascii_digit() {
                            source_index += 1;
                        }
                    }
                }
                if source_index == start {
                    break;
                }
                let token: String = source_chars[start..source_index].iter().collect();
                if let Some(value) = parse_int_token(&token, spec) {
                    if !suppress_assignment {
                        stmts.push(assign_expr(target, lit_int(value)));
                        count += 1;
                    }
                } else {
                    break;
                }
            }
            'f' => {
                skip_source_ws(&mut source_index);
                if source_index >= source_chars.len() {
                    input_failure = true;
                    break;
                }
                let start = source_index;
                while source_index < source_chars.len()
                    && matches!(
                        source_chars[source_index],
                        '0'..='9' | '+' | '-' | '.' | 'e' | 'E'
                    )
                {
                    source_index += 1;
                }
                if source_index == start {
                    break;
                }
                let token: String = source_chars[start..source_index].iter().collect();
                if let Ok(value) = token.trim().parse::<f64>() {
                    if !suppress_assignment {
                        stmts.push(assign_expr(target, lit_float(value)));
                        count += 1;
                    }
                } else {
                    break;
                }
            }
            'c' => {
                if source_index >= source_chars.len() {
                    input_failure = true;
                    break;
                }
                let take = width.unwrap_or(1).max(1);
                let end = source_index.saturating_add(take).min(source_chars.len());
                if take == 1 {
                    let ch = source_chars[source_index];
                    source_index += 1;
                    if !suppress_assignment {
                        stmts.push(assign_expr(target, lit_int(ch as i64)));
                    }
                } else {
                    let token: String = source_chars[source_index..end].iter().collect();
                    source_index = end;
                    if !suppress_assignment {
                        stmts.push(assign_expr(target, lit_str(&token)));
                    }
                }
                if !suppress_assignment {
                    count += 1;
                }
            }
            's' => {
                skip_source_ws(&mut source_index);
                if source_index >= source_chars.len() {
                    input_failure = true;
                    break;
                }
                let start = source_index;
                let reserve_for_next_c = width.is_none()
                    && next_conversion_spec(&format_chars, format_index) == Some('c')
                    && source_index < source_chars.len();
                while source_index < source_chars.len()
                    && !source_chars[source_index].is_whitespace()
                    && (!reserve_for_next_c || source_index + 1 < source_chars.len())
                    && width
                        .map(|limit| source_index - start < limit)
                        .unwrap_or(true)
                {
                    source_index += 1;
                }
                if source_index == start {
                    break;
                }
                let token: String = source_chars[start..source_index].iter().collect();
                if !suppress_assignment {
                    stmts.push(assign_expr(target, lit_str(&token)));
                    count += 1;
                }
            }
            '[' => {
                let mut negate = false;
                if format_index < format_chars.len() && format_chars[format_index] == '^' {
                    negate = true;
                    format_index += 1;
                }
                let mut set_chars = Vec::new();
                while format_index < format_chars.len() && format_chars[format_index] != ']' {
                    set_chars.push(format_chars[format_index]);
                    format_index += 1;
                }
                if format_index < format_chars.len() && format_chars[format_index] == ']' {
                    format_index += 1;
                }
                let set_chars = expand_scan_set(&set_chars);
                let start = source_index;
                if source_index >= source_chars.len() {
                    input_failure = true;
                    break;
                }
                while source_index < source_chars.len() {
                    let ch = source_chars[source_index];
                    let in_set = set_chars.contains(&ch);
                    let should_stop = if negate { in_set } else { !in_set };
                    if should_stop
                        || width
                            .map(|limit| source_index - start >= limit)
                            .unwrap_or(false)
                    {
                        break;
                    }
                    source_index += 1;
                }
                if source_index == start {
                    break;
                }
                let token: String = source_chars[start..source_index].iter().collect();
                if !suppress_assignment {
                    stmts.push(assign_expr(target, lit_str(&token)));
                    count += 1;
                }
            }
            'n' => {
                if !suppress_assignment {
                    stmts.push(assign_expr(target, lit_int(source_index as i64)));
                }
            }
            _ => {}
        }
    }

    stmts.push(lit_int(if count == 0 && input_failure {
        -1
    } else {
        count
    }));
    if stmts.len() == 1 {
        stmts.pop().unwrap()
    } else {
        e(ExprKind::Sequence(stmts))
    }
}

fn next_conversion_spec(format_chars: &[char], mut index: usize) -> Option<char> {
    while index < format_chars.len() {
        if format_chars[index].is_whitespace() {
            index += 1;
            continue;
        }
        if format_chars[index] != '%' {
            return None;
        }
        index += 1;
        if index < format_chars.len() && format_chars[index] == '%' {
            index += 1;
            continue;
        }
        while index < format_chars.len()
            && matches!(format_chars[index], '-' | '+' | ' ' | '#' | '0')
        {
            index += 1;
        }
        while index < format_chars.len() && format_chars[index].is_ascii_digit() {
            index += 1;
        }
        while index < format_chars.len()
            && matches!(format_chars[index], 'l' | 'h' | 'L' | 'z' | 'j' | 't')
        {
            index += 1;
        }
        return format_chars.get(index).copied();
    }
    None
}

fn expand_scan_set(raw: &[char]) -> Vec<char> {
    let mut out = Vec::new();
    let mut i = 0usize;
    while i < raw.len() {
        if i + 2 < raw.len() && raw[i + 1] == '-' {
            let start = raw[i] as u32;
            let end = raw[i + 2] as u32;
            if start <= end {
                for code in start..=end {
                    if let Some(ch) = char::from_u32(code) {
                        out.push(ch);
                    }
                }
                i += 3;
                continue;
            }
        }
        out.push(raw[i]);
        i += 1;
    }
    out
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
                ObjectProperty::KeyValue {
                    key: lit_str("buf"),
                    value: lit_str("") },
                ObjectProperty::KeyValue {
                    key: lit_str("eof"),
                    value: lit_int(0) },
                ObjectProperty::KeyValue {
                    key: lit_str("allow_blocking"),
                    value: lit_int(0) },
            ]))),
            array_bounds: None,
            with_events: false }],
        kind: VarDeclKind::Var }));

    // __libc_stdin_set_blocking(flag): opt-in to real blocking stdin reads.
    out.push(function(
        "__libc_stdin_set_blocking",
        vec!["flag"],
        vec![
            expr_stmt(assign_expr(
                member(stdin_state(), "allow_blocking"),
                ident("flag"),
            )),
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
                bin(
                    BinOp::Eq,
                    member(stdin_state(), "allow_blocking"),
                    lit_int(0),
                ),
                vec![ret(lit_str(""))],
                None,
            ),
            var_decl("__l", call(ident("__libc_readline"), vec![])),
            if_stmt(
                bin(
                    BinOp::Eq,
                    e(ExprKind::Unary {
                        op: UnaryOp::Typeof,
                        expr: Box::new(ident("__l")) }),
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
                            bin(
                                BinOp::LtEq,
                                call(ident("__c_char_code_at"), vec![stdin_buf(), lit_int(0)]),
                                lit_int(32),
                            ),
                        ),
                        vec![expr_stmt(assign_expr(
                            stdin_buf(),
                            call_member(stdin_buf(), "substring", vec![lit_int(1)]),
                        ))],
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
                                    bin(
                                        BinOp::Add,
                                        bin(BinOp::Add, stdin_buf(), ident("line")),
                                        lit_str(" "),
                                    ),
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
                    bin(
                        BinOp::Gt,
                        call(ident("__c_char_code_at"), vec![stdin_buf(), ident("i")]),
                        lit_int(32),
                    ),
                ),
                vec![expr_stmt(assign_expr(
                    ident("i"),
                    bin(BinOp::Add, ident("i"), lit_int(1)),
                ))],
            ),
            var_decl(
                "tok",
                call_member(stdin_buf(), "substring", vec![lit_int(0), ident("i")]),
            ),
            expr_stmt(assign_expr(
                stdin_buf(),
                call_member(stdin_buf(), "substring", vec![ident("i")]),
            )),
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
                        Some(vec![expr_stmt(assign_expr(
                            stdin_buf(),
                            bin(BinOp::Add, ident("line"), lit_str("\n")),
                        ))]),
                    ),
                ],
            ),
            if_stmt(
                bin(BinOp::Eq, member(stdin_buf(), "length"), lit_int(0)),
                vec![ret(lit_str(""))],
                None,
            ),
            var_decl(
                "ch",
                call_member(stdin_buf(), "substring", vec![lit_int(0), lit_int(1)]),
            ),
            expr_stmt(assign_expr(
                stdin_buf(),
                call_member(stdin_buf(), "substring", vec![lit_int(1)]),
            )),
            ret(ident("ch")),
        ],
    ));

    out
}

/// Runtime helper that renders a `char[]` value as a C string for `%s`/`puts`.
/// The storage is polymorphic — a JS string (`strcpy`), a carray pointer (a
/// decayed flexible array `{__ref_kind:"carray",__base,__idx}`), or a plain
/// code-point array (element writes). Strings pass through (up to NUL); the
/// other shapes are decoded one code point at a time with the single-arg
/// `String.fromCharCode` opcode — a runtime-sized spread can't drive the
/// fixed-arity opcode, hence the loop. Stops at the NUL terminator.
///
/// ```text
/// function __libc_char_to_str(v) {
///   if (typeof v === "string") return v.split("\0")[0];
///   var a = v;
///   if (v != null && v.__ref_kind === "carray") a = v.__base.slice(v.__idx);
///   var r = "";
///   var i = 0;
///   while (i < a.length) {
///     var c = a[i];
///     if (c == 0) return r;
///     r = r + String.fromCharCode(c);
///     i = i + 1;
///   }
///   return r;
/// }
/// ```
pub fn char_to_str_runtime_helper() -> Statement {
    let index_a_i = e(ExprKind::Index {
        object: Box::new(ident("a")),
        index: Box::new(ident("i")),
        null_safe: false });
    function(
        "__libc_char_to_str",
        vec!["v"],
        vec![
            // typeof v === "string" → string already; clip at NUL. indexOf +
            // substring, not split — split doesn't resolve in the C runtime.
            if_stmt(
                bin(
                    BinOp::Eq,
                    e(ExprKind::Unary {
                        op: UnaryOp::Typeof,
                        expr: Box::new(ident("v")) }),
                    lit_str("string"),
                ),
                vec![
                    var_decl("k", call_member(ident("v"), "indexOf", vec![lit_str("\0")])),
                    ret(ternary(
                        bin(BinOp::Lt, ident("k"), lit_int(0)),
                        ident("v"),
                        call_member(ident("v"), "substring", vec![lit_int(0), ident("k")]),
                    )),
                ],
                None,
            ),
            var_decl("a", ident("v")),
            // carray pointer → take the slice from __idx onward.
            if_stmt(
                bin(
                    BinOp::And,
                    bin(BinOp::NotEq, ident("v"), e(ExprKind::Lit(Literal::Null))),
                    bin(
                        BinOp::Eq,
                        member(ident("v"), "__ref_kind"),
                        lit_str("carray"),
                    ),
                ),
                vec![expr_stmt(assign_expr(
                    ident("a"),
                    call_member(
                        member(ident("v"), "__base"),
                        "slice",
                        vec![member(ident("v"), "__idx")],
                    ),
                ))],
                None,
            ),
            // Accumulate into a STRING — array push/join don't resolve in the
            // C runtime; string `+` does.
            var_decl("out", lit_str("")),
            var_decl("i", lit_int(0)),
            while_stmt(
                bin(BinOp::Lt, ident("i"), member(ident("a"), "length")),
                vec![
                    var_decl("c", index_a_i),
                    // NUL terminates — and so does reading past a string
                    // base's end (`"abcd"[4]` is undefined where C has '\0').
                    if_stmt(
                        bin(
                            BinOp::Or,
                            bin(BinOp::Eq, ident("c"), lit_int(0)),
                            bin(
                                BinOp::Eq,
                                e(ExprKind::Unary {
                                    op: UnaryOp::Typeof,
                                    expr: Box::new(ident("c")) }),
                                lit_str("undefined"),
                            ),
                        ),
                        vec![ret(ident("out"))],
                        None,
                    ),
                    expr_stmt(assign_expr(
                        ident("out"),
                        bin(
                            BinOp::Add,
                            ident("out"),
                            ternary(
                                bin(
                                    BinOp::Eq,
                                    e(ExprKind::Unary {
                                        op: UnaryOp::Typeof,
                                        expr: Box::new(ident("c")) }),
                                    lit_str("string"),
                                ),
                                ident("c"),
                                call(
                                    member(ident("String"), "fromCharCode"),
                                    vec![ident("c")],
                                ),
                            ),
                        ),
                    )),
                    expr_stmt(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), lit_int(1)),
                    )),
                ],
            ),
            ret(ident("out")),
        ],
    )
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
        let reader = if *spec == 'c' {
            "__libc_stdin_char"
        } else {
            "__libc_stdin_token"
        };
        seq.push(assign_expr(
            ident(&tok_var),
            ternary(ident(&ok_var), call(ident(reader), vec![]), lit_str("")),
        ));
        let converted = match spec {
            'd' | 'u' => bin(
                BinOp::Or,
                call(ident("parseInt"), vec![ident(&tok_var), lit_int(10)]),
                lit_int(0),
            ),
            'i' => bin(
                BinOp::Or,
                call(ident("parseInt"), vec![ident(&tok_var)]),
                lit_int(0),
            ),
            'x' | 'X' => bin(
                BinOp::Or,
                call(ident("parseInt"), vec![ident(&tok_var), lit_int(16)]),
                lit_int(0),
            ),
            'o' => bin(
                BinOp::Or,
                call(ident("parseInt"), vec![ident(&tok_var), lit_int(8)]),
                lit_int(0),
            ),
            'f' | 'e' | 'g' | 'F' | 'E' | 'G' | 'a' => bin(
                BinOp::Or,
                call(ident("parseFloat"), vec![ident(&tok_var)]),
                lit_float(0.0),
            ),
            _ => ident(&tok_var) };
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
