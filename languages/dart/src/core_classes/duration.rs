//! `dart:core`'s `Duration`, as a class.

use super::builders::*;
use vybe_ast::{ClassMember, ExprKind, Expression, InterpolPart, Statement};
use vybe_compiler::primitives::datetime as dt;

/// The Duration components, each with the number of milliseconds it holds.
///
/// **These are FIELDS, not getters, and the names are the ones the previous
/// `wrap_duration_ms` wrapper wrote.** Both choices are deliberate: dart is
/// case-sensitive, so a class field's storage name is the source name
/// unchanged (`js_member_storage_name_for_class` mangles only `#private`), and
/// the DateTime emitters that read `inMilliseconds` off a Duration with a
/// plain `STRUCT_GET` keep working against a class instance with no change.
/// A getter would not be a field and every one of those reads would miss.
/// Spans come from `primitives::datetime`, never spelled here. That module owns
/// the unit vocabulary precisely so `86_400_000` stops appearing in language
/// code — it was counted in 18 places across 8 files before the primitive
/// existed, dart among them.
const DURATION_PARTS: &[(&str, f64)] = &[
    ("inMilliseconds", 1.0),
    ("inMicroseconds", 1.0 / MICROS_PER_MS),
    ("inSeconds", dt::MS_PER_SECOND),
    ("inMinutes", dt::MS_PER_MINUTE),
    ("inHours", dt::MS_PER_HOUR),
    ("inDays", dt::MS_PER_DAY),
];

/// Dart's `Duration` resolves to MICROseconds, one tick finer than the
/// millisecond `primitives::datetime` works in — `EpochPrecision::Micros` in
/// `DateTimePolicy` terms. This is the single conversion at that boundary.
const MICROS_PER_MS: f64 = 1_000.0;

/// `this.inMilliseconds`
fn dur_ms() -> Expression {
    this_field("inMilliseconds")
}

/// `<expr>.inMilliseconds` — the other operand of a binary operator.
fn other_ms() -> Expression {
    field_of(ident("other"), "inMilliseconds")
}

/// `Duration(<ms>)` — construction, so an operator result is a real Duration
/// with the same rtt, vtable and prototype as any other instance.
fn new_duration(ms: Expression) -> Expression {
    Expression::with_span(
        ExprKind::New {
            class: Box::new(ident("Duration")),
            args: vec![vybe_ast::Argument::positional(ms)],
        },
        span(),
    )
}

/// A binary `Duration op Duration` returning a Duration.
fn dur_binop(name: &str, op: vybe_ast::BinOp, rhs: Expression) -> ClassMember {
    method(
        name,
        vec![param("other", Some("Duration"), None)],
        Some("Duration"),
        vec![ret(new_duration(binary(op, dur_ms(), rhs)))],
    )
}

/// `class Duration { … }` — Dart's `dart:core` duration, as a class.
///
/// The wrapper this replaces was an anonymous struct carrying a `__type`
/// STRING of `"Duration"`, which several emitters then string-COMPARED to
/// decide what a value was (`emit_dart_abs`, `emit_num_is_negative`). That is
/// the marker tower [[project_tree_types_have_no_rtt]] describes: a wrapper
/// has no identity, so `d1 + d2` had no `+` to find and fell through to
/// numeric addition on an object — `wasm:js-number.toF64 — not a number`.
/// As a class it fills `ProtocolSlot::Add`/`Sub`/`Mul`/`Neg`/`Compare` through
/// `protocol::canonical_method`, and the shared operator dispatch finds them.
///
/// The walker already collapses `Duration(days: 14, hours: 3)` to a single
/// positional millisecond expression (`normalize_duration_args`), so the
/// constructor needs exactly one optional parameter.
pub(super) fn duration() -> Statement {
    let mut members: Vec<ClassMember> = DURATION_PARTS
        .iter()
        .map(|(name, _)| field(name, "num", num_lit(0.0)))
        .collect();
    // `Duration([num ms = 0])` — every component derived from the one value,
    // exactly as `wrap_duration_ms` derived them.
    let ctor_body: Vec<Statement> = DURATION_PARTS
        .iter()
        .map(|(name, per_ms)| {
            let value = if *per_ms == 1.0 {
                ident("ms")
            } else if *per_ms < 1.0 {
                // microseconds: 1 ms is 1000 of them.
                binary(vybe_ast::BinOp::Mul, ident("ms"), num_lit(1.0 / *per_ms))
            } else {
                binary(vybe_ast::BinOp::IDiv, ident("ms"), num_lit(*per_ms))
            };
            set_this(name, value)
        })
        .collect();
    members.push(constructor(
        vec![param("ms", Some("num"), Some(num_lit(0.0)))],
        ctor_body,
    ));

    members.push(dur_binop("operator+", vybe_ast::BinOp::Add, other_ms()));
    members.push(dur_binop("operator-", vybe_ast::BinOp::Sub, other_ms()));
    // `d * 2` scales by a NUMBER, so the operand is the factor itself.
    members.push(method(
        "operator*",
        vec![param("other", Some("num"), None)],
        Some("Duration"),
        vec![ret(new_duration(binary(
            vybe_ast::BinOp::Mul,
            dur_ms(),
            ident("other"),
        )))],
    ));
    // Unary minus arrives from the walker as `operator-@unary` so it cannot
    // collide with the binary `-` above; built directly, it takes that spelling.
    members.push(method(
        "operator-@unary",
        vec![],
        Some("Duration"),
        vec![ret(new_duration(binary(
            vybe_ast::BinOp::Sub,
            num_lit(0.0),
            dur_ms(),
        )))],
    ));
    // Every member below MUST be declared, even where a receiver-blind
    // `[value_methods]` entry already serves the name.
    //
    // Becoming a class flips `user_typed_receiver_shadow` (`calls.rs:5993`) on
    // for a directly-named Duration local: a receiver whose class IS known
    // deliberately SKIPS the value-method table, because a real type's members
    // are its own. So the moment `Duration` stopped being a wrapper,
    // `d.abs()`, `d.compareTo(other)` and `d.negate()` stopped reaching
    // `common:dart.abs` and friends and answered `undefined` instead. The class
    // has to carry its whole surface, not the half the tests happened to reach
    // through an adapter.
    //
    // `toString` is exempt from that shadow by name at the same site, which is
    // why it worked before these were added.
    //
    // `isNegative` is a zero-arg METHOD rather than the field it looks like:
    // the walker rewrites every `x.isNegative` into a call
    // (`is_dart_zero_arg_getter`) before dispatch sees it, so a field would be
    // read and then invoked — "bool is not callable".
    members.push(method(
        "isNegative",
        vec![],
        Some("bool"),
        vec![ret(binary(vybe_ast::BinOp::Lt, dur_ms(), num_lit(0.0)))],
    ));
    let magnitude = ternary(
        binary(vybe_ast::BinOp::Lt, dur_ms(), num_lit(0.0)),
        binary(vybe_ast::BinOp::Sub, num_lit(0.0), dur_ms()),
        dur_ms(),
    );
    members.push(method(
        "abs",
        vec![],
        Some("Duration"),
        vec![ret(new_duration(magnitude))],
    ));
    // Dart has no `Duration.negate()`; the walker maps unary `-` onto it, so
    // the class answers both spellings with the same body.
    members.push(method(
        "negate",
        vec![],
        Some("Duration"),
        vec![ret(new_duration(binary(
            vybe_ast::BinOp::Sub,
            num_lit(0.0),
            dur_ms(),
        )))],
    ));
    members.push(method(
        "compareTo",
        vec![param("other", Some("Duration"), None)],
        Some("int"),
        vec![ret(ternary(
            binary(vybe_ast::BinOp::Lt, dur_ms(), other_ms()),
            int_lit(-1),
            ternary(
                binary(vybe_ast::BinOp::Gt, dur_ms(), other_ms()),
                int_lit(1),
                int_lit(0),
            ),
        ))],
    ));
    members.push(method(
        "toString",
        vec![],
        Some("String"),
        duration_to_string_body(),
    ));

    class("Duration", members)
}

/// Dart renders a Duration as `H:MM:SS.uuuuuu` — hours unpadded, everything
/// else zero-filled, and a leading `-` for a negative span.
///
/// Spelled as ordinary Dart so it lowers through the shared `~/`, `%` and
/// `padLeft` machinery rather than a Dart-private formatter emit.
fn duration_to_string_body() -> Vec<Statement> {
    // `us` is the magnitude; the sign is re-attached at the end.
    let us = ident("us");
    // The same spans as `DURATION_PARTS`, carried up into MICROseconds — the
    // unit Dart renders at. Derived from `primitives::datetime`, not respelled.
    let micros = |ms: f64| (ms * MICROS_PER_MS) as i64;
    let idiv = |left: Expression, by: i64| binary(vybe_ast::BinOp::IDiv, left, int_lit(by));
    let rem = |left: Expression, by: i64| binary(vybe_ast::BinOp::Mod, left, int_lit(by));
    let pad = |value: Expression, width: i64| {
        call_member(
            call_member(value, "toString", vec![]),
            "padLeft",
            vec![int_lit(width), str_lit("0")],
        )
    };
    vec![
        local(
            "sign",
            ternary(
                binary(
                    vybe_ast::BinOp::Lt,
                    this_field("inMicroseconds"),
                    num_lit(0.0),
                ),
                str_lit("-"),
                str_lit(""),
            ),
        ),
        local(
            "us",
            ternary(
                binary(
                    vybe_ast::BinOp::Lt,
                    this_field("inMicroseconds"),
                    num_lit(0.0),
                ),
                binary(
                    vybe_ast::BinOp::Sub,
                    num_lit(0.0),
                    this_field("inMicroseconds"),
                ),
                this_field("inMicroseconds"),
            ),
        ),
        ret(Expression::with_span(
            ExprKind::Interpolation(vec![
                InterpolPart::Expr(ident("sign")),
                InterpolPart::Expr(idiv(us.clone(), micros(dt::MS_PER_HOUR))),
                InterpolPart::Text(":".to_string()),
                InterpolPart::Expr(pad(rem(idiv(us.clone(), micros(dt::MS_PER_MINUTE)), 60), 2)),
                InterpolPart::Text(":".to_string()),
                InterpolPart::Expr(pad(rem(idiv(us.clone(), micros(dt::MS_PER_SECOND)), 60), 2)),
                InterpolPart::Text(".".to_string()),
                InterpolPart::Expr(pad(rem(us, micros(dt::MS_PER_SECOND)), 6)),
            ]),
            span(),
        )),
    ]
}
