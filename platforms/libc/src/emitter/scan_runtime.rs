//! Formatted input over a runtime character source (C11 7.21.6.2). The format
//! is a compile-time literal: each directive lowers to code over a byte cursor
//! read through `__c_source_byte`, so a string value, a linear address and a
//! carray pointer scan identically. The helpers here delimit one field; the
//! caller converts the field text and stores it.

use crate::emitter::build::*;
use vybe_ast::{BinOp, ExprKind, Expression, Statement, StmtKind};

/// Field shapes `__c_scan_span` delimits.
const KIND_DECIMAL: i64 = 0;
const KIND_HEX: i64 = 1;
const KIND_OCTAL: i64 = 2;
const KIND_AUTO_INTEGER: i64 = 3;
const KIND_FLOAT: i64 = 4;
const KIND_NON_SPACE: i64 = 5;
const KIND_CHARS: i64 = 6;

pub fn runtime_helpers() -> Vec<Statement> {
    vec![
        scan_space(),
        scan_skip(),
        scan_digit(),
        scan_span(),
        scan_set_span(),
        source_slice(),
        scan_auto_integer(),
        scan_copy_bytes(),
    ]
}

/// One conversion specification of the format.
pub struct ScanConversion {
    pub spec: char,
    /// Length modifier spelling: "", "hh", "h", "l", "ll", "j", "z", "t", "L".
    pub length: String,
    pub width: Option<usize>,
    /// Index of the receiving argument after the format.
    pub argument: usize,
}

/// What a store receives for one matched field.
pub struct ScanValue {
    /// The converted value: a number for numeric conversions and `%n`, the
    /// field text for `%s`, `%[` and `%c`.
    pub value: Expression,
    pub source: Expression,
    pub start: Expression,
    pub end: Expression,
}

impl ScanConversion {
    /// The C type a receiving pointer of this conversion points to.
    pub fn target_type(&self) -> &'static str {
        let unsigned = matches!(self.spec, 'u' | 'o' | 'x' | 'X' | 'p');
        match self.spec {
            'f' | 'F' | 'e' | 'E' | 'g' | 'G' | 'a' | 'A' => match self.length.as_str() {
                "l" => "double",
                "L" => "long double",
                _ => "float",
            },
            _ => match (self.length.as_str(), unsigned) {
                ("hh", false) => "signed char",
                ("hh", true) => "unsigned char",
                ("h", false) => "short",
                ("h", true) => "unsigned short",
                ("l", false) => "long",
                ("l", true) => "unsigned long",
                ("ll" | "j" | "q", false) => "long long",
                ("ll" | "j" | "q", true) => "unsigned long long",
                ("z", _) => "size_t",
                ("t", _) => "ptrdiff_t",
                (_, false) => "int",
                (_, true) => "unsigned int",
            },
        }
    }
}

/// `sscanf(source, format, ...)`: the number of assigned fields, or EOF (-1)
/// when the input ends before the first conversion completes.
pub fn sscanf_runtime(
    source: Expression,
    format: &str,
    argument_count: usize,
    tmp_id: u32,
    store: &mut dyn FnMut(&ScanConversion, ScanValue) -> Expression,
) -> Expression {
    let var = |name: &str| ident(&format!("__c_scan_{name}{tmp_id}"));
    let (src, pos, count, done, ok, eof, end) =
        (var("src"), var("pos"), var("n"), var("done"), var("ok"), var("eof"), var("end"));
    let byte_at = |at: Expression| call_expr(ident("__c_source_byte"), vec![src.clone(), at]);
    let increment = |name: &Expression| assign_expr(name.clone(), bin(BinOp::Add, name.clone(), int_lit(1)));
    let input_failure = || {
        sequence(vec![
            assign_expr(
                eof.clone(),
                bin(BinOp::And, bin(BinOp::Eq, byte_at(pos.clone()), int_lit(0)), bin(BinOp::Eq, done.clone(), int_lit(0))),
            ),
            assign_expr(ok.clone(), int_lit(0)),
        ])
    };
    let skip_space = || assign_expr(pos.clone(), call_expr(ident("__c_scan_skip"), vec![src.clone(), pos.clone()]));
    let guarded = |step: Expression| ternary(ok.clone(), step, int_lit(0));
    let match_literal = |byte: u8| {
        guarded(ternary(
            bin(BinOp::Eq, byte_at(pos.clone()), int_lit(i64::from(byte))),
            increment(&pos),
            input_failure(),
        ))
    };

    let mut steps = vec![
        assign_expr(src.clone(), source),
        assign_expr(pos.clone(), int_lit(0)),
        assign_expr(count.clone(), int_lit(0)),
        assign_expr(done.clone(), int_lit(0)),
        assign_expr(ok.clone(), int_lit(1)),
        assign_expr(eof.clone(), int_lit(0)),
    ];
    let bytes = format.as_bytes();
    let mut at = 0usize;
    let mut argument = 0usize;
    while at < bytes.len() {
        let byte = bytes[at];
        if byte.is_ascii_whitespace() {
            while at < bytes.len() && bytes[at].is_ascii_whitespace() {
                at += 1;
            }
            steps.push(guarded(skip_space()));
            continue;
        }
        if byte != b'%' {
            steps.push(match_literal(byte));
            at += 1;
            continue;
        }
        at += 1;
        if bytes.get(at) == Some(&b'%') {
            at += 1;
            steps.push(guarded(skip_space()));
            steps.push(match_literal(b'%'));
            continue;
        }
        let suppress = bytes.get(at) == Some(&b'*');
        if suppress {
            at += 1;
        }
        let digits_start = at;
        while at < bytes.len() && bytes[at].is_ascii_digit() {
            at += 1;
        }
        let width = format[digits_start..at].parse::<usize>().ok().filter(|width| *width > 0);
        let length_start = at;
        while at < bytes.len() && matches!(bytes[at], b'h' | b'l' | b'L' | b'j' | b'z' | b't' | b'q') {
            at += 1;
        }
        let length = format[length_start..at].to_string();
        let Some(&spec_byte) = bytes.get(at) else {
            break;
        };
        at += 1;
        let spec = spec_byte as char;
        let mut set = None;
        if spec == '[' {
            let negate = bytes.get(at) == Some(&b'^');
            if negate {
                at += 1;
            }
            let members_start = at;
            if bytes.get(at) == Some(&b']') {
                at += 1;
            }
            while at < bytes.len() && bytes[at] != b']' {
                at += 1;
            }
            set = Some((scan_set_members(&bytes[members_start..at]), negate));
            at += 1;
        }
        let conversion = ScanConversion { spec, length, width, argument };
        if !suppress {
            if argument >= argument_count {
                break;
            }
            argument += 1;
        }
        if spec == 'n' {
            if !suppress {
                let stored = store(&conversion, ScanValue {
                    value: pos.clone(),
                    source: src.clone(),
                    start: pos.clone(),
                    end: pos.clone(),
                });
                steps.push(guarded(stored));
            }
            continue;
        }
        let kind = match spec {
            'd' | 'u' => KIND_DECIMAL,
            'x' | 'X' | 'p' => KIND_HEX,
            'o' => KIND_OCTAL,
            'i' => KIND_AUTO_INTEGER,
            'f' | 'F' | 'e' | 'E' | 'g' | 'G' | 'a' | 'A' => KIND_FLOAT,
            's' => KIND_NON_SPACE,
            'c' => KIND_CHARS,
            '[' => KIND_CHARS,
            _ => break,
        };
        let field_width = width.unwrap_or(if spec == 'c' { 1 } else { 0 }) as i64;
        let span = match &set {
            Some((members, negate)) => call_expr(
                ident("__c_scan_set_span"),
                vec![src.clone(), pos.clone(), int_lit(field_width), str_lit(members), int_lit(i64::from(*negate))],
            ),
            None => call_expr(
                ident("__c_scan_span"),
                vec![src.clone(), pos.clone(), int_lit(kind), int_lit(field_width)],
            ),
        };
        let text = call_expr(ident("__c_source_slice"), vec![src.clone(), pos.clone(), end.clone()]);
        let value = match spec {
            'd' | 'u' => call_expr(ident("parseInt"), vec![text, int_lit(10)]),
            'x' | 'X' | 'p' => call_expr(ident("parseInt"), vec![text, int_lit(16)]),
            'o' => call_expr(ident("parseInt"), vec![text, int_lit(8)]),
            'i' => call_expr(ident("__c_scan_auto_integer"), vec![text]),
            's' | 'c' | '[' => text,
            _ => call_expr(ident("parseFloat"), vec![text]),
        };
        let mut matched = Vec::new();
        if !suppress {
            matched.push(store(&conversion, ScanValue {
                value,
                source: src.clone(),
                start: pos.clone(),
                end: end.clone(),
            }));
            matched.push(increment(&count));
        }
        matched.push(increment(&done));
        matched.push(assign_expr(pos.clone(), end.clone()));
        let mut conversion_steps = Vec::new();
        if !matches!(spec, 'c' | '[') {
            conversion_steps.push(skip_space());
        }
        conversion_steps.push(ternary(
            bin(BinOp::Eq, byte_at(pos.clone()), int_lit(0)),
            input_failure(),
            sequence(vec![
                assign_expr(end.clone(), span),
                ternary(
                    bin(BinOp::Eq, end.clone(), pos.clone()),
                    assign_expr(ok.clone(), int_lit(0)),
                    sequence(matched),
                ),
            ]),
        ));
        steps.push(guarded(sequence(conversion_steps)));
    }
    steps.push(ternary(eof, int_lit(-1), count));
    sequence(steps)
}

/// The bytes a `%[...]` set admits, ranges `a-z` expanded.
fn scan_set_members(spelling: &[u8]) -> String {
    let mut members = String::new();
    let mut at = 0;
    while at < spelling.len() {
        let first = spelling[at];
        if spelling.get(at + 1) == Some(&b'-') && at + 2 < spelling.len() {
            let last = spelling[at + 2];
            if first <= last {
                members.extend((first..=last).map(char::from));
                at += 3;
                continue;
            }
        }
        members.push(char::from(first));
        at += 1;
    }
    members
}

fn bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary { op, left: Box::new(left), right: Box::new(right) })
}

fn ternary(cond: Expression, then: Expression, else_: Expression) -> Expression {
    expr(ExprKind::Ternary { cond: Box::new(cond), then: Box::new(then), else_: Box::new(else_) })
}

fn sequence(items: Vec<Expression>) -> Expression {
    expr(ExprKind::Sequence(items))
}

fn ret(value: Expression) -> Statement {
    stmt(StmtKind::Return(Some(value)))
}

fn expr_stmt(value: Expression) -> Statement {
    stmt(StmtKind::Expr(value))
}

fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::While { cond, body, else_body: None })
}

fn byte_of(source: &str, index: Expression) -> Expression {
    call_expr(ident("__c_source_byte"), vec![ident(source), index])
}

fn add(name: &str, amount: i64) -> Expression {
    bin(BinOp::Add, ident(name), int_lit(amount))
}

/// `i = i + 1; b = __c_source_byte(src, i)`.
fn advance(index: &str, byte: &str) -> Vec<Statement> {
    vec![
        expr_stmt(assign_expr(ident(index), add(index, 1))),
        expr_stmt(assign_expr(ident(byte), byte_of("src", ident(index)))),
    ]
}

fn below_limit(index: &str) -> Expression {
    bin(BinOp::Lt, ident(index), ident("lim"))
}

fn digit(byte: Expression, base: Expression) -> Expression {
    call_expr(ident("__c_scan_digit"), vec![byte, base])
}

/// `__c_scan_space(b)` — `isspace` in the C locale.
fn scan_space() -> Statement {
    function_stmt(
        "__c_scan_space",
        vec!["b"],
        vec![ret(bin(
            BinOp::Or,
            bin(BinOp::Eq, ident("b"), int_lit(32)),
            bin(BinOp::And, bin(BinOp::GtEq, ident("b"), int_lit(9)), bin(BinOp::LtEq, ident("b"), int_lit(13))),
        ))],
    )
}

/// `__c_scan_skip(src, pos)` — the position after any white space.
fn scan_skip() -> Statement {
    function_stmt(
        "__c_scan_skip",
        vec!["src", "pos"],
        vec![
            while_stmt(
                call_expr(ident("__c_scan_space"), vec![byte_of("src", ident("pos"))]),
                vec![expr_stmt(assign_expr(ident("pos"), add("pos", 1)))],
            ),
            ret(ident("pos")),
        ],
    )
}

/// `__c_scan_digit(b, base)` — whether byte `b` is a digit of `base` (8, 10, 16).
fn scan_digit() -> Statement {
    let between = |low: i64, high: i64, value: Expression| {
        bin(BinOp::And, bin(BinOp::GtEq, value.clone(), int_lit(low)), bin(BinOp::LtEq, value, int_lit(high)))
    };
    let folded = bin(BinOp::BitOr, ident("b"), int_lit(32));
    function_stmt(
        "__c_scan_digit",
        vec!["b", "base"],
        vec![
            if_stmt(bin(BinOp::Eq, ident("base"), int_lit(8)), vec![ret(between(48, 55, ident("b")))], None),
            if_stmt(
                bin(BinOp::Eq, ident("base"), int_lit(16)),
                vec![ret(bin(BinOp::Or, between(48, 57, ident("b")), between(97, 102, folded)))],
                None,
            ),
            ret(between(48, 57, ident("b"))),
        ],
    )
}

/// `__c_scan_span(src, pos, kind, width)` — the end of the longest field of
/// `kind` starting at `pos` within `width` bytes (0 = unbounded), or `pos`
/// when no field matches.
fn scan_span() -> Statement {
    let kind_is = |kind: i64| bin(BinOp::Eq, ident("kind"), int_lit(kind));
    let byte_is = |value: i64| bin(BinOp::Eq, ident("b"), int_lit(value));
    let not_end = bin(BinOp::NotEq, ident("b"), int_lit(0));
    let digits_loop = |index: &str, byte: &str, base: Expression| {
        while_stmt(
            bin(BinOp::And, below_limit(index), digit(ident(byte), base)),
            advance(index, byte),
        )
    };
    let exponent_byte = bin(BinOp::Eq, bin(BinOp::BitOr, ident("b"), int_lit(32)), int_lit(101));
    let signed_c = bin(
        BinOp::Or,
        bin(BinOp::Eq, ident("c"), int_lit(43)),
        bin(BinOp::Eq, ident("c"), int_lit(45)),
    );
    let float_body = vec![
        expr_stmt(assign_expr(ident("start"), ident("i"))),
        digits_loop("i", "b", int_lit(10)),
        expr_stmt(assign_expr(ident("digits"), bin(BinOp::Sub, ident("i"), ident("start")))),
        if_stmt(
            bin(BinOp::And, below_limit("i"), byte_is(46)),
            {
                let mut body = advance("i", "b");
                body.push(expr_stmt(assign_expr(ident("start"), ident("i"))));
                body.push(digits_loop("i", "b", int_lit(10)));
                body.push(expr_stmt(assign_expr(
                    ident("digits"),
                    bin(BinOp::Add, ident("digits"), bin(BinOp::Sub, ident("i"), ident("start"))),
                )));
                body
            },
            None,
        ),
        if_stmt(bin(BinOp::Eq, ident("digits"), int_lit(0)), vec![ret(ident("pos"))], None),
        if_stmt(
            bin(BinOp::And, below_limit("i"), exponent_byte),
            vec![
                expr_stmt(assign_expr(ident("j"), add("i", 1))),
                expr_stmt(assign_expr(ident("c"), byte_of("src", ident("j")))),
                if_stmt(bin(BinOp::And, below_limit("j"), signed_c), advance("j", "c"), None),
                if_stmt(
                    bin(BinOp::And, below_limit("j"), digit(ident("c"), int_lit(10))),
                    vec![
                        digits_loop("j", "c", int_lit(10)),
                        expr_stmt(assign_expr(ident("i"), ident("j"))),
                    ],
                    None,
                ),
            ],
            None,
        ),
        ret(ident("i")),
    ];
    let hex_prefix = bin(
        BinOp::And,
        bin(BinOp::And, byte_is(48), bin(BinOp::Lt, add("i", 2), ident("lim"))),
        bin(
            BinOp::And,
            bin(
                BinOp::Eq,
                bin(BinOp::BitOr, byte_of("src", add("i", 1)), int_lit(32)),
                int_lit(120),
            ),
            digit(byte_of("src", add("i", 2)), int_lit(16)),
        ),
    );
    function_stmt(
        "__c_scan_span",
        vec!["src", "pos", "kind", "width"],
        vec![
            var_decl_stmt("lim", ternary(
                bin(BinOp::Gt, ident("width"), int_lit(0)),
                bin(BinOp::Add, ident("pos"), ident("width")),
                int_lit(2147483647),
            )),
            var_decl_stmt("i", ident("pos")),
            var_decl_stmt("b", byte_of("src", ident("pos"))),
            var_decl_stmt("base", int_lit(10)),
            var_decl_stmt("start", int_lit(0)),
            var_decl_stmt("digits", int_lit(0)),
            var_decl_stmt("j", int_lit(0)),
            var_decl_stmt("c", int_lit(0)),
            if_stmt(
                kind_is(KIND_NON_SPACE),
                vec![
                    while_stmt(
                        bin(
                            BinOp::And,
                            bin(BinOp::And, below_limit("i"), not_end.clone()),
                            expr(ExprKind::Unary {
                                op: vybe_ast::UnaryOp::Not,
                                expr: Box::new(call_expr(ident("__c_scan_space"), vec![ident("b")])),
                            }),
                        ),
                        advance("i", "b"),
                    ),
                    ret(ident("i")),
                ],
                None,
            ),
            if_stmt(
                kind_is(KIND_CHARS),
                vec![
                    while_stmt(bin(BinOp::And, below_limit("i"), not_end), advance("i", "b")),
                    ret(ident("i")),
                ],
                None,
            ),
            if_stmt(
                bin(BinOp::And, below_limit("i"), bin(BinOp::Or, byte_is(43), byte_is(45))),
                advance("i", "b"),
                None,
            ),
            if_stmt(kind_is(KIND_FLOAT), float_body, None),
            if_stmt(kind_is(KIND_HEX), vec![expr_stmt(assign_expr(ident("base"), int_lit(16)))], None),
            if_stmt(kind_is(KIND_OCTAL), vec![expr_stmt(assign_expr(ident("base"), int_lit(8)))], None),
            if_stmt(
                bin(
                    BinOp::And,
                    bin(BinOp::Or, kind_is(KIND_HEX), kind_is(KIND_AUTO_INTEGER)),
                    hex_prefix,
                ),
                vec![
                    expr_stmt(assign_expr(ident("i"), add("i", 2))),
                    expr_stmt(assign_expr(ident("b"), byte_of("src", ident("i")))),
                    expr_stmt(assign_expr(ident("base"), int_lit(16))),
                ],
                Some(vec![if_stmt(
                    bin(BinOp::And, kind_is(KIND_AUTO_INTEGER), byte_is(48)),
                    vec![expr_stmt(assign_expr(ident("base"), int_lit(8)))],
                    None,
                )]),
            ),
            expr_stmt(assign_expr(ident("start"), ident("i"))),
            digits_loop("i", "b", ident("base")),
            if_stmt(bin(BinOp::Eq, ident("i"), ident("start")), vec![ret(ident("pos"))], None),
            ret(ident("i")),
        ],
    )
}

/// `__c_scan_set_span(src, pos, width, set, negate)` — the end of the run of
/// bytes in (or, negated, not in) `set`, within `width` bytes (0 = unbounded).
fn scan_set_span() -> Statement {
    let member = call_expr(ident("__c_strtok_is_delimiter"), vec![ident("b"), ident("set")]);
    function_stmt(
        "__c_scan_set_span",
        vec!["src", "pos", "width", "set", "negate"],
        vec![
            var_decl_stmt("lim", ternary(
                bin(BinOp::Gt, ident("width"), int_lit(0)),
                bin(BinOp::Add, ident("pos"), ident("width")),
                int_lit(2147483647),
            )),
            var_decl_stmt("i", ident("pos")),
            var_decl_stmt("b", byte_of("src", ident("pos"))),
            while_stmt(
                bin(
                    BinOp::And,
                    bin(BinOp::And, below_limit("i"), bin(BinOp::NotEq, ident("b"), int_lit(0))),
                    bin(BinOp::NotEq, ternary(member, int_lit(1), int_lit(0)), ident("negate")),
                ),
                advance("i", "b"),
            ),
            ret(ident("i")),
        ],
    )
}

/// `__c_source_slice(src, start, end)` — bytes `[start, end)` of a character
/// source as a string value.
fn source_slice() -> Statement {
    function_stmt(
        "__c_source_slice",
        vec!["src", "start", "end"],
        vec![
            var_decl_stmt("text", str_lit("")),
            while_stmt(
                bin(BinOp::Lt, ident("start"), ident("end")),
                vec![
                    expr_stmt(assign_expr(
                        ident("text"),
                        bin(
                            BinOp::Concat,
                            ident("text"),
                            call_member(ident("String"), "fromCharCode", vec![byte_of("src", ident("start"))]),
                        ),
                    )),
                    expr_stmt(assign_expr(ident("start"), add("start", 1))),
                ],
            ),
            ret(ident("text")),
        ],
    )
}

/// `__c_scan_auto_integer(text)` — `%i`: the base follows the prefix, `0x`
/// hexadecimal, a leading `0` octal, otherwise decimal.
fn scan_auto_integer() -> Statement {
    let code = |at: Expression| call_expr(ident("__c_char_code_at"), vec![ident("text"), at]);
    function_stmt(
        "__c_scan_auto_integer",
        vec!["text"],
        vec![
            var_decl_stmt("k", int_lit(0)),
            var_decl_stmt("c", code(int_lit(0))),
            if_stmt(
                bin(BinOp::Or, bin(BinOp::Eq, ident("c"), int_lit(43)), bin(BinOp::Eq, ident("c"), int_lit(45))),
                vec![expr_stmt(assign_expr(ident("k"), int_lit(1)))],
                None,
            ),
            if_stmt(
                bin(
                    BinOp::And,
                    bin(BinOp::Eq, code(ident("k")), int_lit(48)),
                    bin(BinOp::Gt, member(ident("text"), "length"), add("k", 1)),
                ),
                vec![if_stmt(
                    bin(
                        BinOp::NotEq,
                        bin(BinOp::BitOr, code(add("k", 1)), int_lit(32)),
                        int_lit(120),
                    ),
                    vec![ret(call_expr(ident("parseInt"), vec![ident("text"), int_lit(8)]))],
                    None,
                )],
                None,
            ),
            ret(call_expr(ident("parseInt"), vec![ident("text")])),
        ],
    )
}

/// `__c_scan_copy_bytes(dst, src, start, end)` — `%c`: the field bytes into
/// linear `dst`, with no terminator.
fn scan_copy_bytes() -> Statement {
    function_stmt(
        "__c_scan_copy_bytes",
        vec!["dst", "src", "start", "end"],
        vec![
            var_decl_stmt("k", int_lit(0)),
            while_stmt(
                bin(BinOp::Lt, ident("start"), ident("end")),
                vec![
                    expr_stmt(call_expr(
                        ident("__c_ptr_i32_store8"),
                        vec![bin(BinOp::Add, ident("dst"), ident("k")), byte_of("src", ident("start"))],
                    )),
                    expr_stmt(assign_expr(ident("k"), add("k", 1))),
                    expr_stmt(assign_expr(ident("start"), add("start", 1))),
                ],
            ),
            ret(ident("dst")),
        ],
    )
}
