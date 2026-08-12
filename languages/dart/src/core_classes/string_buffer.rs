//! `dart:core`'s `StringBuffer`, as a class.

use super::builders::*;
use vybe_ast::{Expression, Statement, StmtKind};

/// The backing field. Underscore-prefixed so it reads as private in Dart and
/// cannot collide with a user member on a subclass.
const BUF: &str = "_vybeBuf";

/// `this.<BUF>`
fn buf() -> Expression {
    this_field(BUF)
}

/// `this.<BUF> = <value>;`
fn set_buf(value: Expression) -> Statement {
    assign(buf(), value)
}

/// `class StringBuffer { … }` — Dart's `dart:core` buffer, as a class.
pub(super) fn string_buffer() -> Statement {
    let members = vec![
        field(BUF, "String", str_lit("")),
        // `StringBuffer([Object? content = ''])` — the walker supplies the
        // omitted-argument default as an explicit `""`. Do NOT put the default
        // here: the shared optional-param lowering uses null as its missing
        // sentinel, but Dart distinguishes omitted from explicit null, and
        // `StringBuffer(null)` must contain `"null"`.
        constructor(
            vec![param("content", None, None)],
            vec![set_buf(stringify(ident("content")))],
        ),
        // **Getters, and that is the RIGHT shape** — Dart spells them
        // `int get length`, and a getter is what publishes the PROTOCOL SLOT:
        // `normalize_class.rs` maps the property name through
        // `protocol::canonical_method`, so `length` records `Len` in
        // `special_methods` and the class stamps it on its prototype.
        //
        // Declaring them as METHODS instead is what is catastrophic: it puts
        // `length` into the FLAT, class-less `defined_class_methods` set
        // (`calls.rs:5995`) that every untyped receiver consults, so `got.length`
        // on a plain String in the test harness itself was diverted to a
        // StringBuffer member. **MEASURED: every dart slice went to 0 passing —
        // 0/50, 0/56, 0/57, 0/44.**
        //
        // The BODY is a member READ. There is no `[value_methods] length` row
        // to call any more — `length` is a `Property` leaf on this type in the
        // tree — and `_vybeBuf` holds a plain String, whose `.length` the
        // shared member read already answers.
        getter("length", "int", field_of(buf(), "length")),
        getter("isEmpty", "bool", call_member(buf(), "isEmpty", vec![])),
        getter(
            "isNotEmpty",
            "bool",
            call_member(buf(), "isNotEmpty", vec![]),
        ),
        method(
            "write",
            vec![param("o", None, None)],
            Some("void"),
            vec![set_buf(concat(buf(), stringify(ident("o"))))],
        ),
        method(
            "writeln",
            vec![param("o", None, None)],
            Some("void"),
            vec![set_buf(concat(
                concat(buf(), stringify(ident("o"))),
                str_lit("\n"),
            ))],
        ),
        method(
            "writeAll",
            vec![
                param("objects", None, None),
                param("separator", Some("String"), Some(str_lit(""))),
            ],
            Some("void"),
            vec![set_buf(concat(
                buf(),
                call_member(ident("objects"), "join", vec![ident("separator")]),
            ))],
        ),
        method(
            "writeCharCode",
            vec![param("charCode", Some("int"), None)],
            Some("void"),
            vec![set_buf(concat(
                buf(),
                call_member(ident("String"), "fromCharCode", vec![ident("charCode")]),
            ))],
        ),
        method("clear", vec![], Some("void"), vec![set_buf(str_lit(""))]),
        // The whole point: a real `toString` member, so member dispatch finds
        // it on the receiver instead of reaching `Object.prototype.toString`.
        // `normalize_class` binds it to `ProtocolSlot::ToString` from here.
        method(
            "toString",
            vec![],
            Some("String"),
            vec![Statement::with_span(StmtKind::Return(Some(buf())), span())],
        ),
    ];
    class("StringBuffer", members)
}
