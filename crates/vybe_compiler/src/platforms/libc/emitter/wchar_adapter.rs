//! C wide-character (`wchar.h`) adapter — WASM/libc-faithful model.
//!
//! A `wchar_t[]` is a flat, NUL-terminated array of UTF-32 code points, exactly
//! as clang→wasm32-wasi lays it out in linear memory (wasi-libc sets
//! `wchar_t = int`, 4 bytes). It is **not** a JS string: core WASM has no string
//! value type (see the spec — `utf8` appears only for `name`s / text literals),
//! and the js-string-builtins surface is immutable with no addressing, so it
//! cannot back C pointer arithmetic. Wide buffers therefore use the same flat
//! array + `carray` pointer model as `int[]`/`int*`.
//!
//! Conversion to/from a JS string happens only at boundaries (formatting,
//! display) — the analogue of the proposal's `fromCharCodeArray` /
//! `intoCharCodeArray`.

use crate::ast::{
    Argument, BinOp, BindingPattern, ExprKind, Expression, Literal, Modifiers, Param, PassBy,
    Statement, StmtKind, VarDeclKind, VarDeclarator,
};
use crate::platforms::libc::emitter::pointers;

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}
fn s(kind: StmtKind) -> Statement {
    Statement::new(kind)
}
fn ident(n: &str) -> Expression {
    e(ExprKind::Ident(n.to_string()))
}
fn lit_int(n: i64) -> Expression {
    e(ExprKind::Lit(Literal::Int(n)))
}
fn lit_str(v: &str) -> Expression {
    e(ExprKind::Lit(Literal::Str(v.to_string())))
}
fn null_lit() -> Expression {
    e(ExprKind::Lit(Literal::Null))
}
fn member(o: Expression, f: &str) -> Expression {
    e(ExprKind::Member {
        object: Box::new(o),
        field: f.to_string(),
        null_safe: false,
    })
}
fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    e(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}
fn call_member(o: Expression, f: &str, args: Vec<Expression>) -> Expression {
    call(member(o, f), args)
}
fn bin(op: BinOp, l: Expression, r: Expression) -> Expression {
    e(ExprKind::Binary {
        op,
        left: Box::new(l),
        right: Box::new(r),
    })
}
fn index(o: Expression, i: Expression) -> Expression {
    e(ExprKind::Index {
        object: Box::new(o),
        index: Box::new(i),
        null_safe: false,
    })
}
fn assign(target: Expression, value: Expression) -> Expression {
    e(ExprKind::Assign {
        target: Box::new(target),
        value: Box::new(value),
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
fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    s(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
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
        else_body,
    })
}
fn ret(v: Expression) -> Statement {
    s(StmtKind::Return(Some(v)))
}
fn expr_stmt(v: Expression) -> Statement {
    s(StmtKind::Expr(v))
}
fn function(name: &str, params: Vec<&str>, body: Vec<Statement>) -> Statement {
    s(StmtKind::FunctionDecl {
        name: name.to_string(),
        params: params
            .into_iter()
            .map(|p| Param {
                name: p.to_string(),
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

/// `L"hello"` → flat NUL-terminated code-point array `[104,101,108,108,111,0]`.
pub fn wide_string_literal(text: &str) -> Expression {
    crate::platforms::libc::emitter::arrays::carray_from_string_literal(text)
}

// Wide buffers are `int`-typed flat arrays (wchar_t == int on wasm32-wasi), so
// array-method dispatch (`.indexOf`/`.slice`/`.length`) works inline on the typed
// value. Only the runtime string→array conversion needs a typed helper.

/// NUL index of a wide array, guarded: `i = arr.indexOf(0); i < 0 ? arr.length : i`.
fn nul_index(arr: Expression) -> Expression {
    let idx = call_member(arr.clone(), "indexOf", vec![lit_int(0)]);
    e(ExprKind::Ternary {
        cond: Box::new(bin(BinOp::Lt, idx.clone(), lit_int(0))),
        then: Box::new(member(arr, "length")),
        else_: Box::new(idx),
    })
}

/// `wcslen(s)`: number of wide chars up to the NUL.
pub fn wcslen(arr: Expression) -> Expression {
    nul_index(arr)
}

/// `wcscpy(dst, src)`: copy src code points through the NUL into dst, return dst.
/// `dst = src.slice(0, wcslen(src) + 1)`.
pub fn wcscpy(dst: Expression, src: Expression) -> Expression {
    let end = bin(BinOp::Add, nul_index(src.clone()), lit_int(1));
    assign(dst, call_member(src, "slice", vec![lit_int(0), end]))
}

/// `wcscmp(a, b)`: lexicographic compare → -1 / 0 / 1, via the boundary string
/// forms (code points compare identically).
pub fn wcscmp(a: Expression, b: Expression) -> Expression {
    let sa = wide_to_string(a);
    let sb = wide_to_string(b);
    e(ExprKind::Ternary {
        cond: Box::new(bin(BinOp::Lt, sa.clone(), sb.clone())),
        then: Box::new(lit_int(-1)),
        else_: Box::new(e(ExprKind::Ternary {
            cond: Box::new(bin(BinOp::Gt, sa, sb)),
            then: Box::new(lit_int(1)),
            else_: Box::new(lit_int(0)),
        })),
    })
}

/// Convert a wide array to a JS string up to the NUL (boundary conversion).
/// Compile-time fold for a literal; otherwise inline `String.fromCharCode(...)`
/// over the array slice (the array is int-typed, so the spread/slice dispatch).
pub fn wide_to_string(arr: Expression) -> Expression {
    if let Some(text) = literal_wide_to_string(&arr) {
        return lit_str(&text);
    }
    let slice = call_member(arr.clone(), "slice", vec![lit_int(0), nul_index(arr)]);
    e(ExprKind::Call {
        callee: Box::new(member(ident("String"), "fromCharCode")),
        args: vec![Argument {
            value: slice,
            name: None,
            by_ref: false,
            spread: true,
        }],
        optional: false,
    })
}

/// Convert a runtime JS string into a NUL-terminated wide code-point array.
pub fn string_to_wide(str_expr: Expression) -> Expression {
    call(ident("__libc_str_to_wide"), vec![str_expr])
}

/// If `arr` is a literal code-point array (as produced by `wide_string_literal`),
/// reconstruct the Rust string up to the NUL — lets wprintf/swprintf reuse the
/// narrow sprintf format parser without a runtime conversion.
pub fn literal_wide_to_string(arr: &Expression) -> Option<String> {
    let ExprKind::Array(elems) = &arr.kind else {
        return None;
    };
    let mut out = String::new();
    for el in elems {
        let ExprKind::Lit(Literal::Int(code)) = &el.value.kind else {
            return None;
        };
        if *code == 0 {
            break;
        }
        out.push(char::from_u32(*code as u32)?);
    }
    Some(out)
}

fn lt(l: Expression, r: Expression) -> Expression {
    bin(BinOp::Lt, l, r)
}
fn lte(l: Expression, r: Expression) -> Expression {
    bin(BinOp::LtEq, l, r)
}
fn eq(l: Expression, r: Expression) -> Expression {
    bin(BinOp::Eq, l, r)
}
fn ne(l: Expression, r: Expression) -> Expression {
    bin(BinOp::NotEq, l, r)
}
fn and(l: Expression, r: Expression) -> Expression {
    bin(BinOp::And, l, r)
}
fn incr(name: &str) -> Statement {
    expr_stmt(assign(
        ident(name),
        bin(BinOp::Add, ident(name), lit_int(1)),
    ))
}

fn min_expr(a: Expression, b: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(bin(BinOp::Lt, a.clone(), b.clone())),
        then: Box::new(a),
        else_: Box::new(b),
    })
}

fn ptr_or_null(base: Expression, idx: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(bin(BinOp::Lt, idx.clone(), lit_int(0))),
        then: Box::new(null_lit()),
        else_: Box::new(pointers::make_carray_ptr(base, idx)),
    })
}

pub fn wcsnlen(arr: Expression, n: Expression) -> Expression {
    min_expr(nul_index(arr), n)
}

pub fn wcsncmp(a: Expression, b: Expression, n: Expression) -> Expression {
    wcscmp(
        call_member(a, "slice", vec![lit_int(0), n.clone()]),
        call_member(b, "slice", vec![lit_int(0), n]),
    )
}

pub fn wcschr(arr: Expression, ch: Expression) -> Expression {
    let searchable = call_member(
        arr.clone(),
        "slice",
        vec![
            lit_int(0),
            bin(BinOp::Add, nul_index(arr.clone()), lit_int(1)),
        ],
    );
    let idx = call_member(searchable, "indexOf", vec![ch]);
    ptr_or_null(arr, idx)
}

pub fn wcsrchr(arr: Expression, ch: Expression) -> Expression {
    let searchable = call_member(
        arr.clone(),
        "slice",
        vec![
            lit_int(0),
            bin(BinOp::Add, nul_index(arr.clone()), lit_int(1)),
        ],
    );
    let idx = call_member(searchable, "lastIndexOf", vec![ch]);
    ptr_or_null(arr, idx)
}

pub fn wcsstr(hay: Expression, needle: Expression) -> Expression {
    let idx = call_member(
        wide_to_string(hay.clone()),
        "indexOf",
        vec![wide_to_string(needle)],
    );
    ptr_or_null(hay, idx)
}

pub fn wcspbrk(arr: Expression, accept: Expression) -> Expression {
    let idx = call(ident("__libc_wcspbrk_idx"), vec![arr.clone(), accept]);
    ptr_or_null(arr, idx)
}

pub fn wcsspn(arr: Expression, accept: Expression) -> Expression {
    call(ident("__libc_wcsspn"), vec![arr, accept])
}

pub fn wcscspn(arr: Expression, reject: Expression) -> Expression {
    call(ident("__libc_wcscspn"), vec![arr, reject])
}

pub fn wmemcmp(a: Expression, b: Expression, n: Expression) -> Expression {
    call(ident("__libc_wmemcmp"), vec![a, b, n])
}

pub fn wmemchr(arr: Expression, ch: Expression, n: Expression) -> Expression {
    let idx = call_member(
        call_member(arr.clone(), "slice", vec![lit_int(0), n]),
        "indexOf",
        vec![ch],
    );
    ptr_or_null(arr, idx)
}

pub fn wcsncpy(dst: Expression, src: Expression, n: Expression) -> Expression {
    call(ident("__libc_wcsncpy"), vec![dst, src, n])
}

pub fn wcscat(dst: Expression, src: Expression) -> Expression {
    call(ident("__libc_wcscat"), vec![dst, src])
}

pub fn wcsncat(dst: Expression, src: Expression, n: Expression) -> Expression {
    call(ident("__libc_wcsncat"), vec![dst, src, n])
}

pub fn wmemcpy(dst: Expression, src: Expression, n: Expression) -> Expression {
    call(ident("__libc_wmemcpy"), vec![dst, src, n])
}

pub fn wmemset(dst: Expression, ch: Expression, n: Expression) -> Expression {
    call(ident("__libc_wmemset"), vec![dst, ch, n])
}

pub fn wcsdup(src: Expression) -> Expression {
    call_member(src, "slice", vec![lit_int(0)])
}

/// Runtime boundary helper: convert a runtime JS string into a NUL-terminated
/// code-point array. The `s` param is string-typed so `.length`/`.charCodeAt`
/// dispatch as string ops; the result array is built with `.push`.
pub fn runtime_helpers() -> Vec<Statement> {
    // Build the FunctionDecl directly so the string param carries a type hint.
    let body = vec![
        var_decl("a", e(ExprKind::Array(Vec::new()))),
        var_decl("i", lit_int(0)),
        while_stmt(
            lt(ident("i"), member(ident("s"), "length")),
            vec![
                expr_stmt(call_member(
                    ident("a"),
                    "push",
                    vec![call(
                        ident("__c_char_code_at"),
                        vec![ident("s"), ident("i")],
                    )],
                )),
                incr("i"),
            ],
        ),
        expr_stmt(call_member(ident("a"), "push", vec![lit_int(0)])),
        ret(ident("a")),
    ];
    let mut decl = function("__libc_str_to_wide", vec!["s"], body);
    if let StmtKind::FunctionDecl { params, .. } = &mut decl.kind {
        if let Some(p) = params.first_mut() {
            p.type_hint = Some("char*".to_string());
        }
    }
    let wcsncpy_body = vec![
        var_decl("i", lit_int(0)),
        var_decl("len", nul_index(ident("src"))),
        while_stmt(
            lt(ident("i"), ident("n")),
            vec![
                if_stmt(
                    lt(ident("i"), ident("len")),
                    vec![expr_stmt(assign(
                        index(ident("dst"), ident("i")),
                        index(ident("src"), ident("i")),
                    ))],
                    Some(vec![expr_stmt(assign(
                        index(ident("dst"), ident("i")),
                        lit_int(0),
                    ))]),
                ),
                incr("i"),
            ],
        ),
        ret(ident("dst")),
    ];

    let wcscat_body = vec![
        var_decl("start", nul_index(ident("dst"))),
        var_decl("i", lit_int(0)),
        var_decl("len", nul_index(ident("src"))),
        while_stmt(
            lte(ident("i"), ident("len")),
            vec![
                expr_stmt(assign(
                    index(ident("dst"), bin(BinOp::Add, ident("start"), ident("i"))),
                    index(ident("src"), ident("i")),
                )),
                incr("i"),
            ],
        ),
        ret(ident("dst")),
    ];

    let wcsncat_body = vec![
        var_decl("start", nul_index(ident("dst"))),
        var_decl("i", lit_int(0)),
        while_stmt(
            and(
                lt(ident("i"), ident("n")),
                ne(index(ident("src"), ident("i")), lit_int(0)),
            ),
            vec![
                expr_stmt(assign(
                    index(ident("dst"), bin(BinOp::Add, ident("start"), ident("i"))),
                    index(ident("src"), ident("i")),
                )),
                incr("i"),
            ],
        ),
        expr_stmt(assign(
            index(ident("dst"), bin(BinOp::Add, ident("start"), ident("i"))),
            lit_int(0),
        )),
        ret(ident("dst")),
    ];

    let wmemcpy_body = vec![
        var_decl("i", lit_int(0)),
        while_stmt(
            lt(ident("i"), ident("n")),
            vec![
                expr_stmt(assign(
                    index(ident("dst"), ident("i")),
                    index(ident("src"), ident("i")),
                )),
                incr("i"),
            ],
        ),
        ret(ident("dst")),
    ];

    let wmemset_body = vec![
        var_decl("i", lit_int(0)),
        while_stmt(
            lt(ident("i"), ident("n")),
            vec![
                expr_stmt(assign(index(ident("dst"), ident("i")), ident("ch"))),
                incr("i"),
            ],
        ),
        ret(ident("dst")),
    ];

    let wmemcmp_body = vec![
        var_decl("i", lit_int(0)),
        while_stmt(
            lt(ident("i"), ident("n")),
            vec![
                if_stmt(
                    ne(index(ident("a"), ident("i")), index(ident("b"), ident("i"))),
                    vec![ret(e(ExprKind::Ternary {
                        cond: Box::new(lt(
                            index(ident("a"), ident("i")),
                            index(ident("b"), ident("i")),
                        )),
                        then: Box::new(lit_int(-1)),
                        else_: Box::new(lit_int(1)),
                    }))],
                    None,
                ),
                incr("i"),
            ],
        ),
        ret(lit_int(0)),
    ];

    let wcspbrk_body = vec![
        var_decl("i", lit_int(0)),
        while_stmt(
            ne(index(ident("s"), ident("i")), lit_int(0)),
            vec![
                if_stmt(
                    bin(
                        BinOp::GtEq,
                        call_member(
                            ident("accept"),
                            "indexOf",
                            vec![index(ident("s"), ident("i"))],
                        ),
                        lit_int(0),
                    ),
                    vec![ret(ident("i"))],
                    None,
                ),
                incr("i"),
            ],
        ),
        ret(lit_int(-1)),
    ];

    let wcsspn_body = vec![
        var_decl("i", lit_int(0)),
        while_stmt(
            and(
                ne(index(ident("s"), ident("i")), lit_int(0)),
                bin(
                    BinOp::GtEq,
                    call_member(
                        ident("accept"),
                        "indexOf",
                        vec![index(ident("s"), ident("i"))],
                    ),
                    lit_int(0),
                ),
            ),
            vec![incr("i")],
        ),
        ret(ident("i")),
    ];

    let wcscspn_body = vec![
        var_decl("i", lit_int(0)),
        while_stmt(
            and(
                ne(index(ident("s"), ident("i")), lit_int(0)),
                bin(
                    BinOp::Lt,
                    call_member(
                        ident("reject"),
                        "indexOf",
                        vec![index(ident("s"), ident("i"))],
                    ),
                    lit_int(0),
                ),
            ),
            vec![incr("i")],
        ),
        ret(ident("i")),
    ];

    vec![
        decl,
        function("__libc_wcsncpy", vec!["dst", "src", "n"], wcsncpy_body),
        function("__libc_wcscat", vec!["dst", "src"], wcscat_body),
        function("__libc_wcsncat", vec!["dst", "src", "n"], wcsncat_body),
        function("__libc_wmemcpy", vec!["dst", "src", "n"], wmemcpy_body),
        function("__libc_wmemset", vec!["dst", "ch", "n"], wmemset_body),
        function("__libc_wmemcmp", vec!["a", "b", "n"], wmemcmp_body),
        function("__libc_wcspbrk_idx", vec!["s", "accept"], wcspbrk_body),
        function("__libc_wcsspn", vec!["s", "accept"], wcsspn_body),
        function("__libc_wcscspn", vec!["s", "reject"], wcscspn_body),
    ]
}
