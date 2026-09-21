//! POSIX regex.h — `regcomp` / `regexec` / `regfree` / `regerror`, plus the
//! `regex_t` fields programs read (`re_nsub`).
//!
//! Built on the ECMA RegExp surface (`new RegExp(src, flags)` + `.exec`, which
//! returns a match with `.index` and—under the `d` flag—`.indices[i] = [so, eo]`,
//! exactly the offsets POSIX `regmatch_t {rm_so, rm_eo}` needs). The compiled
//! pattern (`regex_t`) is modeled as
//! `{__src, __flags, __nosub, re_nsub, __err}`; `regmatch_t` as `{rm_so, rm_eo}`.
//!
//! Implemented surface:
//!   - `regcomp` validates the pattern at compile time (invalid → `REG_BADPAT`),
//!     records the capture-group count in `re_nsub`, and honors `REG_EXTENDED`
//!     (ECMA regex ≈ POSIX ERE), `REG_ICASE`, `REG_NEWLINE`, `REG_NOSUB`.
//!   - `regexec` fills `pmatch` from the per-group match offsets, returns 0 /
//!     `REG_NOMATCH`, and skips `pmatch` when the pattern was compiled `REG_NOSUB`.
//!     `eflags` is threaded through.
//!   - `regerror` returns the standard message text for an error code.
//!   - `regfree` releases (nothing to release in this model).
//!
//! Shared by any libc-targeting front-end.

use crate::emitter::build::*;
use vybe_ast::{
    BinOp, CatchClause, ExprKind, Expression, Literal, ObjectProperty, Statement, StmtKind,
};

// POSIX error codes (glibc values) used here.
const REG_NOMATCH: i64 = 1;
const REG_BADPAT: i64 = 2;
const REG_ESPACE: i64 = 12;

fn bin(op: BinOp, l: Expression, r: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(l),
        right: Box::new(r),
    })
}

fn obj(pairs: Vec<(&str, Expression)>) -> Expression {
    expr(ExprKind::Object(
        pairs
            .into_iter()
            .map(|(k, v)| ObjectProperty::KeyValue {
                key: str_lit(k),
                value: v,
            })
            .collect(),
    ))
}

fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
}

fn ternary(cond: Expression, then: Expression, else_: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then),
        else_: Box::new(else_),
    })
}

fn new_regexp(src: Expression, flags: Expression) -> Expression {
    call_expr(ident("__c_regexp_new"), vec![src, flags])
}

fn literal_str(expr: &Expression) -> Option<&str> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    }
}

fn literal_int(expr: &Expression) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(v)) => Some(*v),
        ExprKind::Lit(Literal::Float(v)) => Some(*v as i64),
        _ => None,
    }
}

fn posix_pattern_to_ecma(src: &str, cflags: i64) -> String {
    let mut out = src
        .replace("[[:alpha:]]", "[A-Za-z]")
        .replace("[[:digit:]]", "[0-9]")
        .replace("[[:alnum:]]", "[A-Za-z0-9]")
        .replace("[[:space:]]", r"\s");
    if cflags & 1 == 0 {
        let mut converted = String::with_capacity(out.len());
        let mut chars = out.chars().peekable();
        while let Some(ch) = chars.next() {
            if ch == '\\' {
                match chars.peek().copied() {
                    Some('(') | Some(')') => {
                        converted.push(chars.next().unwrap());
                    }
                    _ => {
                        converted.push(ch);
                    }
                }
            } else if matches!(ch, '(' | ')' | '+' | '?' | '{' | '|' | '}') {
                converted.push('\\');
                converted.push(ch);
            } else {
                converted.push(ch);
            }
        }
        out = converted;
    }
    out
}

fn count_ecma_capture_groups(src: &str, cflags: i64) -> i64 {
    if cflags & 8 != 0 {
        return 0;
    }
    let mut count = 0;
    let mut escaped = false;
    let mut in_class = false;
    let chars: Vec<char> = src.chars().collect();
    for (i, ch) in chars.iter().copied().enumerate() {
        if escaped {
            escaped = false;
            continue;
        }
        if ch == '\\' {
            escaped = true;
            continue;
        }
        if ch == '[' {
            in_class = true;
            continue;
        }
        if ch == ']' {
            in_class = false;
            continue;
        }
        if in_class || ch != '(' {
            continue;
        }
        if chars.get(i + 1) == Some(&'?') {
            continue;
        }
        count += 1;
    }
    count
}

fn simple_compile_error(src: &str) -> bool {
    let mut escaped = false;
    let mut in_class = false;
    let mut class_closed = true;
    let mut parens = 0i64;
    for ch in src.chars() {
        if escaped {
            escaped = false;
            continue;
        }
        if ch == '\\' {
            escaped = true;
            continue;
        }
        if in_class {
            if ch == ']' {
                in_class = false;
                class_closed = true;
            }
            continue;
        }
        match ch {
            '[' => {
                in_class = true;
                class_closed = false;
            }
            '(' => parens += 1,
            ')' => {
                parens -= 1;
                if parens < 0 {
                    return true;
                }
            }
            _ => {}
        }
    }
    !class_closed || parens != 0
}

fn literal_regex_t(src: &str, cflags: i64) -> Expression {
    let translated = posix_pattern_to_ecma(src, cflags);
    let err = literal_regex_error_code_from_translated(&translated);
    let mut flags = String::new();
    if cflags & 2 != 0 {
        flags.push('i');
    }
    if cflags & 4 != 0 {
        flags.push('m');
    }
    obj(vec![
        ("__src", str_lit(&translated)),
        ("__flags", str_lit(&flags)),
        ("__nosub", int_lit(if cflags & 8 != 0 { 1 } else { 0 })),
        ("re_nsub", int_lit(count_ecma_capture_groups(&translated, cflags))),
        ("__err", int_lit(err)),
    ])
}

fn literal_regex_error_code_from_translated(translated: &str) -> i64 {
    if simple_compile_error(translated) {
        REG_BADPAT
    } else {
        0
    }
}

fn literal_regex_error_code(src: &str, cflags: i64) -> i64 {
    let translated = posix_pattern_to_ecma(src, cflags);
    literal_regex_error_code_from_translated(&translated)
}

/// `(flags & bit) != 0`
fn flag_set(var: &str, bit: i64) -> Expression {
    bin(
        BinOp::NotEq,
        bin(BinOp::BitAnd, ident(var), int_lit(bit)),
        int_lit(0),
    )
}

// ── call-site lowerings ──────────────────────────────────────────────────────

/// `regcomp(&preg, pattern, cflags)` → compile into `preg`; returns 0 on success
/// or an error code (`REG_BADPAT` for an invalid pattern).
pub fn regcomp(preg_lval: Expression, pattern: Expression, cflags: Expression) -> Expression {
    if let (Some(src), Some(flags)) = (literal_str(&pattern), literal_int(&cflags)) {
        let store = assign_expr(preg_lval.clone(), literal_regex_t(src, flags));
        return expr(ExprKind::Sequence(vec![
            store,
            int_lit(literal_regex_error_code(src, flags)),
        ]));
    }
    let store = assign_expr(
        preg_lval.clone(),
        call_expr(ident("__c_regcomp_compile"), vec![pattern, cflags]),
    );
    expr(ExprKind::Sequence(vec![store, member(preg_lval, "__err")]))
}

/// `regexec(&preg, str, nmatch, pmatch, eflags)` → 0 on match (filling
/// `pmatch` unless compiled `REG_NOSUB`), `REG_NOMATCH` otherwise.
pub fn regexec(
    preg_val: Expression,
    input: Expression,
    nmatch: Expression,
    pmatch: Expression,
    eflags: Expression,
) -> Expression {
    call_expr(
        ident("__c_regexec"),
        vec![preg_val, input, nmatch, pmatch, eflags],
    )
}

/// `regfree(&preg)` → nothing to release in this model.
pub fn regfree() -> Expression {
    int_lit(0)
}

/// `regerror(errcode, &preg, errbuf, errbuf_size)` → write the standard message
/// into `errbuf`, return its length + 1 (size including the NUL, per POSIX).
pub fn regerror(errcode: Expression, errbuf_lval: Expression) -> Expression {
    let store = assign_expr(
        errbuf_lval.clone(),
        call_expr(ident("__c_regerror_msg"), vec![errcode]),
    );
    expr(ExprKind::Sequence(vec![
        store,
        bin(BinOp::Add, member(errbuf_lval, "length"), int_lit(1)),
    ]))
}

// ── runtime helpers (injected once into the program prelude) ─────────────────

pub fn runtime_helpers() -> Vec<Statement> {
    vec![
        regcomp_compile_helper(),
        regexec_helper(),
        regerror_msg_helper(),
    ]
}

/// `__c_regcomp_compile(pat, cflags)` → the `regex_t` object. Maps POSIX compile
/// flags onto ECMA RegExp flags, validates the pattern (invalid → `REG_BADPAT`),
/// and counts capture groups into `re_nsub` (via the `pat|` empty-alternative
/// trick: matching `""` yields one slot per group).
fn regcomp_compile_helper() -> Statement {
    let set_flag = |bit: i64, ch: &str| {
        if_stmt(
            flag_set("cflags", bit),
            vec![stmt(StmtKind::Expr(assign_expr(
                ident("flags"),
                bin(BinOp::Concat, ident("flags"), str_lit(ch)),
            )))],
            None,
        )
    };
    // try { var g = new RegExp(pat + "|", flags); nsub = g.exec("").length - 1; }
    // catch (e) { err = REG_BADPAT; }
    let probe = new_regexp(bin(BinOp::Concat, ident("pat"), str_lit("|")), ident("flags"));
    let try_body = vec![
        var_decl_stmt("g", probe),
        stmt(StmtKind::Expr(assign_expr(
            ident("nsub"),
            bin(
                BinOp::Sub,
                member(call_expr(ident("__c_regexp_exec"), vec![ident("g"), str_lit("")]), "length"),
                int_lit(1),
            ),
        ))),
    ];
    let catch = CatchClause {
        types: Vec::new(),
        var_name: Some("e".to_string()),
        stack_var: None,
        body: vec![stmt(StmtKind::Expr(assign_expr(
            ident("err"),
            int_lit(REG_BADPAT),
        )))],
        when_clause: None,
    };
    function_stmt(
        "__c_regcomp_compile",
        vec!["pat", "cflags"],
        vec![
            var_decl_stmt("flags", str_lit("")),
            set_flag(2, "i"), // REG_ICASE
            set_flag(4, "m"), // REG_NEWLINE
            var_decl_stmt(
                "nosub",
                ternary(flag_set("cflags", 8), int_lit(1), int_lit(0)),
            ), // REG_NOSUB
            var_decl_stmt("nsub", int_lit(0)),
            var_decl_stmt("err", int_lit(0)),
            stmt(StmtKind::Try {
                body: try_body,
                catches: vec![catch],
                else_body: None,
                finally: None,
            }),
            stmt(StmtKind::Return(Some(obj(vec![
                ("__src", ident("pat")),
                ("__flags", ident("flags")),
                ("__nosub", ident("nosub")),
                ("re_nsub", ident("nsub")),
                ("__err", ident("err")),
            ])))),
        ],
    )
}

/// `__c_regexec(preg, str, nmatch, pmatch, eflags)` — run the regex, fill the
/// first `nmatch` `pmatch` entries from the match's group indices
/// (`-1`/`-1` for a group that did not participate), and return 0; return
/// `REG_NOMATCH` when there is no match. `pmatch` is skipped when the pattern
/// was compiled `REG_NOSUB`.
fn regexec_helper() -> Statement {
    let new_re = new_regexp(
        member(ident("preg"), "__src"),
        bin(BinOp::Concat, member(ident("preg"), "__flags"), str_lit("d")),
    );
    let notbol_anchor = bin(
        BinOp::And,
        flag_set("eflags", 1),
        bin(
            BinOp::Eq,
            call_member(member(ident("preg"), "__src"), "indexOf", vec![str_lit("^")]),
            int_lit(0),
        ),
    );
    let noteol_anchor = bin(
        BinOp::And,
        flag_set("eflags", 2),
        call_member(member(ident("preg"), "__src"), "endsWith", vec![str_lit("$")]),
    );
    let sp = index_expr(member(ident("m"), "indices"), ident("i"));
    let fill = if_stmt(
        bin(BinOp::Eq, ident("sp"), null_lit()),
        vec![stmt(StmtKind::Expr(assign_expr(
            index_expr(ident("pmatch"), ident("i")),
            obj(vec![("rm_so", int_lit(-1)), ("rm_eo", int_lit(-1))]),
        )))],
        Some(vec![stmt(StmtKind::Expr(assign_expr(
            index_expr(ident("pmatch"), ident("i")),
            obj(vec![
                ("rm_so", index_expr(ident("sp"), int_lit(0))),
                ("rm_eo", index_expr(ident("sp"), int_lit(1))),
            ]),
        )))]),
    );
    let fill_loop = while_stmt(
        bin(BinOp::Lt, ident("i"), ident("nmatch")),
        vec![
            var_decl_stmt("sp", sp),
            fill,
            stmt(StmtKind::Expr(assign_expr(
                ident("i"),
                bin(BinOp::Add, ident("i"), int_lit(1)),
            ))),
        ],
    );
    function_stmt(
        "__c_regexec",
        vec!["preg", "str", "nmatch", "pmatch", "eflags"],
        vec![
            if_stmt(
                bin(BinOp::Or, notbol_anchor, noteol_anchor),
                vec![stmt(StmtKind::Return(Some(int_lit(REG_NOMATCH))))],
                None,
            ),
            var_decl_stmt("re", new_re),
            var_decl_stmt(
                "m",
                call_expr(ident("__c_regexp_exec"), vec![ident("re"), ident("str")]),
            ),
            if_stmt(
                bin(BinOp::Eq, ident("m"), null_lit()),
                vec![stmt(StmtKind::Return(Some(int_lit(REG_NOMATCH))))],
                None,
            ),
            // REG_NOSUB: don't report sub-expression offsets.
            if_stmt(
                bin(BinOp::Eq, member(ident("preg"), "__nosub"), int_lit(0)),
                vec![var_decl_stmt("i", int_lit(0)), fill_loop],
                None,
            ),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    )
}

/// `__c_regerror_msg(errcode)` → the standard POSIX message text.
fn regerror_msg_helper() -> Statement {
    let msg = ternary(
        bin(BinOp::Eq, ident("errcode"), int_lit(REG_NOMATCH)),
        str_lit("No match"),
        ternary(
            bin(BinOp::Eq, ident("errcode"), int_lit(REG_BADPAT)),
            str_lit("Invalid regular expression"),
            ternary(
                bin(BinOp::Eq, ident("errcode"), int_lit(REG_ESPACE)),
                str_lit("Out of memory"),
                str_lit("Unknown regex error"),
            ),
        ),
    );
    function_stmt(
        "__c_regerror_msg",
        vec!["errcode"],
        vec![stmt(StmtKind::Return(Some(msg)))],
    )
}
