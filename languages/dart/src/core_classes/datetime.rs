//! `dart:core`'s `DateTime`, as a class.
//!
//! # What it replaces
//!
//! `wrap_datetime_ms` (`emitter/core_adapter.rs`) built an anonymous
//! `ecma:object` and stamped a `__type` STRING of `"DateTime"` on it. That
//! string was the type: `emit_dart_add` READ it to decide whether `x.add(y)`
//! meant "advance a moment" or "push onto a list", and `emit_slot_is_type` had
//! to union an rtt test with a `__type` compare precisely because Duration had
//! become a class while DateTime had not. A wrapper is not a type
//! ([[project_tree_types_have_no_rtt]]); with both as classes, `dt.add(d)`
//! resolves on the RECEIVER and the list branch serves only lists.
//!
//! # The calendar parts are FIELDS, and that is forced
//!
//! `year`/`month`/`day`/`hour`/`minute`/`second`/`weekday` are derived, so a
//! getter is what they want to be — and a getter is exactly what cannot work:
//! **a dart property getter's body cannot see `this`**
//! ([[project_dart_property_getter_body_loses_this]]), so every one of them
//! would answer `undefined`. They are eager fields, filled once in the
//! constructor, the same shape `wrap_datetime_ms` produced.
//!
//! Eager is safe because a Dart `DateTime` is IMMUTABLE — `add`, `subtract`
//! and `difference` each build a new instance rather than mutating one — so a
//! part cannot go stale relative to the epoch it was derived from.
//!
//! Do not "fix" these into getters.
//!
//! # `millisecondsSinceEpoch` is the state
//!
//! Everything else is derived from it, and it keeps its source spelling as the
//! storage name so the emitters that read it off a receiver with a plain
//! `STRUCT_GET` keep working — the same contract `Duration.inMilliseconds`
//! holds. `primitives::datetime` owns the unit vocabulary; `MS_PER_*` is never
//! respelled here.

use super::builders::*;
use vybe_ast::{ClassMember, ExprKind, Expression, InterpolPart, Statement};

/// The epoch field. Dart spells it `millisecondsSinceEpoch` and reads it
/// directly, so the storage name is the source name.
const EPOCH: &str = "millisecondsSinceEpoch";

/// Each derived calendar part and the builtin that reads it off an epoch
/// millisecond value.
///
/// Every one of these is a bare `host:ecma:date:getUTC*` row in the profile —
/// no adapter code at all. `month` and `weekday` are absent because neither is
/// a bare read: Dart numbers months from 1 and weekdays Monday=1, which are
/// `MonthIndexing` and `WeekdayBase` in `vybe_ast::datetime`, not arithmetic to
/// respell here.
const PARTS: &[(&str, &str)] = &[
    ("year", "__dart_date_year"),
    ("day", "__dart_date_day"),
    ("hour", "__dart_date_hour"),
    ("minute", "__dart_date_minute"),
    ("second", "__dart_date_second"),
    ("millisecond", "__dart_date_millisecond"),
];

fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(ident(name)),
            args: args
                .into_iter()
                .map(vybe_ast::Argument::positional)
                .collect(),
            optional: false },
        span(),
    )
}

/// `this.millisecondsSinceEpoch`
fn epoch() -> Expression {
    this_field(EPOCH)
}

/// `<other>.millisecondsSinceEpoch`
fn other_epoch() -> Expression {
    field_of(ident("other"), EPOCH)
}

/// `DateTime(<ms>, <utc>)` — a plain construction.
///
/// This went through a top-level trampoline for one day, because a constructor
/// called inside an instance method of its own class adopted the ambient
/// receiver instead of allocating. That was `expressions.rs` saving and
/// restoring `__js_this` around the `New` call without ever clearing it; the
/// clear landed, so `new` allocates unconditionally and the workaround is gone.
fn new_datetime(ms: Expression, utc: Expression) -> Expression {
    Expression::with_span(
        ExprKind::New {
            class: Box::new(ident("DateTime")),
            args: vec![
                vybe_ast::Argument::positional(ms),
                vybe_ast::Argument::positional(utc),
            ] },
        span(),
    )
}

/// `Duration(<ms>)` — `difference` answers a real Duration, not a number.
fn new_duration(ms: Expression) -> Expression {
    Expression::with_span(
        ExprKind::New {
            class: Box::new(ident("Duration")),
            args: vec![vybe_ast::Argument::positional(ms)] },
        span(),
    )
}

/// `<n>.toString().padLeft(<width>, '0')`
fn pad(value: Expression, width: i64) -> Expression {
    call_member(
        call_member(value, "toString", vec![]),
        "padLeft",
        vec![int_lit(width), str_lit("0")],
    )
}

/// `class DateTime { … }`.
pub(super) fn datetime() -> Statement {
    let mut members: Vec<ClassMember> = vec![
        field(EPOCH, "num", num_lit(0.0)),
        field("isUtc", "bool", bool_lit(false)),
        field("month", "int", int_lit(1)),
        field("weekday", "int", int_lit(1)),
    ];
    members.extend(PARTS.iter().map(|(name, _)| field(name, "int", int_lit(0))));

    // `DateTime(num ms, [bool utc = false])`.
    //
    // The walker collapses every source spelling to this one shape before it
    // arrives: `DateTime(2024, 3, 15)` becomes a single epoch expression
    // (`normalize_datetime_args`), `DateTime.now()` and `DateTime.utc(...)`
    // become constructions of this class. One constructor, one representation.
    let mut body: Vec<Statement> = vec![
        set_this(EPOCH, ident("ms")),
        set_this("isUtc", ident("utc")),
    ];
    for (name, reader) in PARTS {
        body.push(set_this(name, call(reader, vec![ident("ms")])));
    }
    // `month` and `weekday` are the two CONVENTIONS, and both are read through
    // a builtin that applies the shared primitive rather than arithmetic
    // spelled here: `dart.date_month` adds the `MonthIndexing::OneBased` offset
    // and `dart.date_weekday` applies `WeekdayBase::MondayOne`.
    body.push(set_this("month", call("__dart_date_month", vec![ident("ms")])));
    body.push(set_this(
        "weekday",
        call("__dart_date_weekday", vec![ident("ms")]),
    ));
    members.push(constructor(
        vec![
            param("ms", Some("num"), Some(num_lit(0.0))),
            param("utc", Some("bool"), Some(bool_lit(false))),
        ],
        body,
    ));

    members.extend(instance_members());
    class("DateTime", members)
}

fn instance_members() -> Vec<ClassMember> {
    let mut out = Vec::new();

    // `add` / `subtract` take a Duration and answer a new moment.
    //
    // Declaring `add` is what retires `emit_dart_add`'s `__type == "DateTime"`
    // sniff. It is safe on the flat, class-less `defined_class_methods` set
    // (`calls.rs:5995`) — measured 2026-08-09: a user class declaring `add`
    // does NOT divert `list.add(x)`, because a list receiver is CLASSIFIED
    // (`builtin_type_of`) and array dispatch answers before the flat set is
    // consulted. That is the difference from `Duration.compareTo`, whose
    // competing receivers were WRAPPERS with no classification at all.
    for (name, op) in [("add", vybe_ast::BinOp::Add), ("subtract", vybe_ast::BinOp::Sub)] {
        out.push(method(
            name,
            vec![param("other", Some("Duration"), None)],
            Some("DateTime"),
            vec![ret(new_datetime(
                binary(op, epoch(), field_of(ident("other"), "inMilliseconds")),
                this_field("isUtc"),
            ))],
        ));
    }

    // `difference` answers a real `Duration`, which is a class — so the span
    // arrives with `inDays`/`inHours`/operators already on it.
    out.push(method(
        "difference",
        vec![param("other", Some("DateTime"), None)],
        Some("Duration"),
        vec![ret(new_duration(binary(
            vybe_ast::BinOp::Sub,
            epoch(),
            other_epoch(),
        )))],
    ));

    for (name, op) in [
        ("isBefore", vybe_ast::BinOp::Lt),
        ("isAfter", vybe_ast::BinOp::Gt),
        ("isAtSameMomentAs", vybe_ast::BinOp::Eq),
    ] {
        out.push(method(
            name,
            vec![param("other", Some("DateTime"), None)],
            Some("bool"),
            vec![ret(binary(op, epoch(), other_epoch()))],
        ));
    }

    // **`compareTo` IS declared here, and that is new.**
    //
    // `Duration` deliberately omits it because the flat method set diverted
    // `someDateTime.compareTo(x)` and `someBigInt.compareTo(x)` away from
    // `common:dart.compare_to` — measured −2 datetime, −4 bigint. A DateTime
    // that is a real class answers `compareTo` from its own body, so the
    // datetime half of that cost is gone. BigInt is a PRIMITIVE
    // ([[project_dart_bigint_is_a_primitive_not_a_class]]) and still has no
    // class to answer from, so Duration's `compareTo` stays blocked on the
    // remaining −4 until the flat set goes (flexclassplan §3a).
    out.push(method(
        "compareTo",
        vec![param("other", Some("DateTime"), None)],
        Some("int"),
        vec![ret(ternary(
            binary(vybe_ast::BinOp::Lt, epoch(), other_epoch()),
            int_lit(-1),
            ternary(
                binary(vybe_ast::BinOp::Gt, epoch(), other_epoch()),
                int_lit(1),
                int_lit(0),
            ),
        ))],
    ));

    out.push(method(
        "toUtc",
        vec![],
        Some("DateTime"),
        vec![ret(new_datetime(epoch(), bool_lit(true)))],
    ));
    out.push(method(
        "toLocal",
        vec![],
        Some("DateTime"),
        vec![ret(new_datetime(epoch(), bool_lit(false)))],
    ));

    out.push(iso8601());
    out.push(method(
        "toString",
        vec![],
        Some("String"),
        vec![ret(rendered(false))],
    ));
    out
}

/// `toIso8601String()` — `YYYY-MM-DDTHH:MM:SS.mmm`, with a `Z` when UTC.
fn iso8601() -> ClassMember {
    method(
        "toIso8601String",
        vec![],
        Some("String"),
        vec![ret(rendered(true))],
    )
}

/// The shared rendering. Dart's `toString` separates date and time with a
/// space; `toIso8601String` uses `T` and appends `Z` for a UTC moment.
fn rendered(iso: bool) -> Expression {
    let sep = if iso { "T" } else { " " };
    let body = interp(vec![
        InterpolPart::Expr(pad(this_field("year"), 4)),
        InterpolPart::Text("-".to_string()),
        InterpolPart::Expr(pad(this_field("month"), 2)),
        InterpolPart::Text("-".to_string()),
        InterpolPart::Expr(pad(this_field("day"), 2)),
        InterpolPart::Text(sep.to_string()),
        InterpolPart::Expr(pad(this_field("hour"), 2)),
        InterpolPart::Text(":".to_string()),
        InterpolPart::Expr(pad(this_field("minute"), 2)),
        InterpolPart::Text(":".to_string()),
        InterpolPart::Expr(pad(this_field("second"), 2)),
        InterpolPart::Text(".".to_string()),
        InterpolPart::Expr(pad(this_field("millisecond"), 3)),
    ]);
    if !iso {
        return body;
    }
    concat(
        body,
        ternary(this_field("isUtc"), str_lit("Z"), str_lit("")),
    )
}

fn interp(parts: Vec<InterpolPart>) -> Expression {
    Expression::with_span(ExprKind::Interpolation(parts), span())
}
