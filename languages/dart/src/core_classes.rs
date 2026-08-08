//! `dart:core` classes the runtime provides, synthesized as ordinary AST.
//!
//! A builtin class is a CLASS. It is declared here as a `StmtKind::ClassDecl`
//! and appended to the module body, so it flows through the same path a user
//! class does — `normalize_class` → `NormalClass` → `compile_class` — and
//! inherits every piece of machinery that path already provides: a reserved
//! type slot and a real rtt at `struct.new_default $T`, the runtime
//! `TypeRegistry` registration, the prototype stamp that makes member dispatch
//! RECEIVER-based, MRO, and protocol-slot binding (flexclassplan §2b's
//! "explicit registration from the frontend").
//!
//! That is the difference from the `[builtins]` entry this replaces. A builtin
//! emitted an anonymous `struct.new 0` carrying a private marker field: no
//! type, no vtable, no prototype, so `sb.toString()` had nothing to dispatch
//! to and fell through to `Object.prototype.toString` → `[object Object]`.
//! Registering the name in the namespace tree fixes RESOLUTION, not identity —
//! namespaceplan §246 records that the write-side rtt migration never happened
//! and that 184 sites stamp a `__type` STRING instead. A normalized class is
//! the one shape that does not add a 185th.
//!
//! Declaring it AS AST rather than as source text is what keeps it free: no
//! scan of the program text, no second parse. Pascal's
//! `synthesize_exception_classes` is the same move.
//!
//! Bodies are deliberately plain: `+`, interpolation, `.join`, `.length`. Each
//! lowers through the shared string/collection machinery, so the class carries
//! no Dart-private buffer emitter and the semantics are the ones every other
//! language gets.

use vybe_ast::{
    ClassMember, ClassModifiers, ExprKind, Expression, InterpolPart, Literal, Modifiers, Param,
    PassBy, PropertySetter, Span, Statement, StmtKind, Visibility };
use vybe_compiler::primitives::datetime as dt;

/// Every `dart:core` class the walker declares, paired with its builder. The
/// walker skips any name the program declares itself, so a user
/// `class StringBuffer` still wins.
///
/// This is also what `tree_register.rs` reads to declare the same names in the
/// namespace tree, so the tree and the AST can never disagree about which
/// classes exist.
pub const CORE_CLASSES: &[(&str, fn() -> Statement)] =
    &[("StringBuffer", string_buffer), ("Duration", duration)];

/// True when `name` is one of the classes above.
///
/// The walker needs this during the WALK — `dart_call_or_new` must know that
/// `StringBuffer(...)` is a construction, and the class it will append does
/// not exist yet at that point. It cannot ask the namespace tree instead:
/// tree registration happens at compile time, after the walk.
pub fn is_core_class(name: &str) -> bool {
    CORE_CLASSES.iter().any(|(n, _)| *n == name)
}

/// The backing field. Underscore-prefixed so it reads as private in Dart and
/// cannot collide with a user member on a subclass.
const BUF: &str = "_vybeBuf";

fn span() -> Span {
    Span::default()
}

fn ident(name: &str) -> Expression {
    Expression::with_span(ExprKind::Ident(name.to_string()), span())
}

fn str_lit(value: &str) -> Expression {
    Expression::with_span(ExprKind::Lit(Literal::Str(value.to_string())), span())
}

/// `this.<BUF>`
fn buf() -> Expression {
    Expression::with_span(
        ExprKind::Member {
            object: Box::new(Expression::with_span(ExprKind::This, span())),
            field: BUF.to_string(),
            null_safe: false },
        span(),
    )
}

/// `this.<BUF> = <value>;`
fn set_buf(value: Expression) -> Statement {
    Statement::with_span(
        StmtKind::Assign {
            targets: vec![buf()],
            value,
            by_ref: false },
        span(),
    )
}

/// `<left> <op> <right>`
fn binary(op: vybe_ast::BinOp, left: Expression, right: Expression) -> Expression {
    Expression::with_span(
        ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right) },
        span(),
    )
}

/// `<left> + <right>`
fn concat(left: Expression, right: Expression) -> Expression {
    binary(vybe_ast::BinOp::Add, left, right)
}

/// `this.<name>`
fn this_field(name: &str) -> Expression {
    Expression::with_span(
        ExprKind::Member {
            object: Box::new(Expression::with_span(ExprKind::This, span())),
            field: name.to_string(),
            null_safe: false },
        span(),
    )
}

/// `<object>.<field>`
fn field_of(object: Expression, name: &str) -> Expression {
    Expression::with_span(
        ExprKind::Member {
            object: Box::new(object),
            field: name.to_string(),
            null_safe: false },
        span(),
    )
}

/// `this.<name> = <value>;`
fn set_this(name: &str, value: Expression) -> Statement {
    Statement::with_span(
        StmtKind::Assign {
            targets: vec![this_field(name)],
            value,
            by_ref: false },
        span(),
    )
}

/// `return <value>;`
fn ret(value: Expression) -> Statement {
    Statement::with_span(StmtKind::Return(Some(value)), span())
}

fn num_lit(value: f64) -> Expression {
    Expression::with_span(ExprKind::Lit(Literal::Float(value)), span())
}

fn int_lit(value: i64) -> Expression {
    Expression::with_span(ExprKind::Lit(Literal::Int(value)), span())
}

/// `<cond> ? <then> : <else_>`
fn ternary(cond: Expression, then: Expression, else_: Expression) -> Expression {
    Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then),
            else_: Box::new(else_) },
        span(),
    )
}

/// `'$expr'` — stringification through the shared `to_string` slot, which is
/// what interpolation lowers to. A `.toString()` call would instead re-enter
/// member dispatch on an arbitrary receiver.
fn stringify(expr: Expression) -> Expression {
    Expression::with_span(
        ExprKind::Interpolation(vec![InterpolPart::Expr(expr)]),
        span(),
    )
}

/// `<object>.<method>(<args>)`
fn call_member(object: Expression, method: &str, args: Vec<Expression>) -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(object),
                    field: method.to_string(),
                    null_safe: false },
                span(),
            )),
            args: args.into_iter().map(vybe_ast::Argument::positional).collect(),
            optional: false },
        span(),
    )
}

fn param(name: &str, type_hint: Option<&str>, default: Option<Expression>) -> Param {
    Param {
        name: name.to_string(),
        type_hint: type_hint.map(|t| t.to_string().into()),
        is_optional: default.is_some(),
        default,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_nullable: false }
}

/// An instance method, as the `FunctionDecl` statement `ClassMember::Method`
/// wraps — the same shape a walker produces for a source-declared method.
fn method(name: &str, params: Vec<Param>, return_type: Option<&str>, body: Vec<Statement>) -> ClassMember {
    ClassMember::Method(Box::new(Statement::with_span(
        StmtKind::FunctionDecl {
            name: name.to_string(),
            params,
            body,
            return_type: return_type.map(str::to_string),
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false },
        span(),
    )))
}

/// Member names the walker rewrites from a property read into a zero-arg CALL
/// (`is_dart_zero_arg_getter`, `walker.rs:6455`) before dispatch ever sees them.
///
/// **A core class must declare any of these as a zero-arg METHOD, never as a
/// getter or a field.** Dart spells them `bool get isEmpty`, but by the time the
/// class is dispatched against, `sb.isEmpty` is already `sb.isEmpty()` — reading
/// a field or property there yields the VALUE and then invokes it:
/// "bool is not callable (type: true)". Measured on both `StringBuffer.isEmpty`
/// and `Duration.isNegative`.
///
/// This list is the checklist for a new core class; the walker owns the real one.
/// A name leaves it only when its ROLE is consumed on the shared member-read
/// path — not because the name reads like a property.
#[allow(dead_code)]
const ZERO_ARG_GETTER_NAMES: &[&str] = &[
    "length", "isEmpty", "isNotEmpty", "isEven", "isOdd", "isNegative", "isNaN", "isFinite", "isInfinite",
    "sign", "first", "last", "single", "singleOrNull", "runes", "codeUnits", "keys",
    "values", "entries", "reversed", "isRunning", "elapsed", "elapsedMilliseconds",
    "elapsedMicroseconds",
];

/// A read-only property — Dart's `int get length => …`.
///
/// A name in [`ZERO_ARG_GETTER_NAMES`] does NOT work as a getter — but see the
/// StringBuffer members for why the method form can be worse.
fn getter(name: &str, type_hint: &str, value: Expression) -> ClassMember {
    ClassMember::Property {
        name: name.to_string(),
        type_hint: Some(type_hint.to_string()),
        getter: Some(vec![Statement::with_span(
            StmtKind::Return(Some(value)),
            span(),
        )]),
        setter: None::<PropertySetter>,
        is_auto: false,
        modifiers: Modifiers::default() }
}

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
            args: vec![vybe_ast::Argument::positional(ms)] },
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
pub fn duration() -> Statement {
    let mut members: Vec<ClassMember> = DURATION_PARTS
        .iter()
        .map(|(name, _)| ClassMember::Field {
            name: name.to_string(),
            type_hint: Some("num".to_string()),
            init: Some(num_lit(0.0)),
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None })
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
                binary(vybe_ast::BinOp::Div, ident("ms"), num_lit(*per_ms))
            };
            set_this(name, value)
        })
        .collect();
    members.push(ClassMember::Constructor {
        name: None,
        params: vec![param("ms", Some("num"), Some(num_lit(0.0)))],
        body: ctor_body,
        base_args: None,
        initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
        visibility: Visibility::Public });

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
    members.push(method("abs", vec![], Some("Duration"), vec![ret(new_duration(magnitude))]));
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
    // **`compareTo` is deliberately absent, and `d.compareTo(other)` is broken
    // because of it — a KNOWN, MEASURED trade, not an oversight.**
    //
    // Two shared mechanisms pull in opposite directions and a core class is
    // caught between them:
    //   - declare it, and `defined_class_methods` (`calls.rs:5995`) — a FLAT,
    //     class-less set of every method name ANY class declares — diverts
    //     `someDateTime.compareTo(x)` and `someBigInt.compareTo(x)` away from
    //     their adapters into a member those types do not have. Measured:
    //     −2 `datetime_compare_*`, −4 `bigint_compare_to_*`.
    //   - omit it, and `user_typed_receiver_shadow` (`calls.rs:5993`) makes a
    //     directly-named Duration local skip `[value_methods]` entirely, so
    //     `common:dart.compare_to` is never reached. Measured: −4
    //     `duration_compare_*`.
    // Omitting costs 4 and declaring costs 6, so it stays out until the flat
    // set goes — the site itself calls it "actively wrong" and slates it for
    // deletion once receiver typing covers untyped locals (flexclassplan §3a).
    //
    // `abs` and `negate` above are safe only because no OTHER receiver in the
    // suite calls them on an untyped local.
    members.push(method(
        "toString",
        vec![],
        Some("String"),
        duration_to_string_body(),
    ));

    Statement::with_span(
        StmtKind::ClassDecl {
            name: "Duration".to_string(),
            parents: Vec::new(),
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
            decorators: vec![] },
        span(),
    )
}

/// Dart renders a Duration as `H:MM:SS.uuuuuu` — hours unpadded, everything
/// else zero-filled, and a leading `-` for a negative span.
///
/// Spelled as ordinary Dart so it lowers through the shared `~/`, `%` and
/// `padLeft` machinery rather than a Dart-private formatter emit.
fn duration_to_string_body() -> Vec<Statement> {
    let local = |name: &str, value: Expression| {
        Statement::with_span(
            StmtKind::VarDecl {
                declarations: vec![vybe_ast::VarDeclarator {
                    pattern: vybe_ast::BindingPattern::Ident(name.to_string()),
                    type_hint: None,
                    init: Some(value),
                    array_bounds: None,
                    with_events: false }],
                kind: vybe_ast::VarDeclKind::Var },
            span(),
        )
    };
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
                binary(vybe_ast::BinOp::Lt, this_field("inMicroseconds"), num_lit(0.0)),
                str_lit("-"),
                str_lit(""),
            ),
        ),
        local(
            "us",
            ternary(
                binary(vybe_ast::BinOp::Lt, this_field("inMicroseconds"), num_lit(0.0)),
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

/// `class StringBuffer { … }` — Dart's `dart:core` buffer, as a class.
pub fn string_buffer() -> Statement {
    let members = vec![
        ClassMember::Field {
            name: BUF.to_string(),
            type_hint: Some("String".to_string()),
            init: Some(str_lit("")),
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None },
        // `StringBuffer([Object content = ''])` — the optional seed.
        ClassMember::Constructor {
            name: None,
            params: vec![param("content", Some("Object"), Some(str_lit("")))],
            body: vec![set_buf(stringify(ident("content")))],
            base_args: None,
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public },
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
        // The BODY is an explicit call because a spliced expression never meets
        // `is_dart_zero_arg_getter` — that rewrite runs over parsed source. The
        // call is what carries `_buf.length` to `emit_dart_length`, the slot
        // consumer; a bare member read here is a `STRUCT_GET "length"` on a
        // String, which no slot answers.
        getter("length", "int", call_member(buf(), "length", vec![])),
        getter("isEmpty", "bool", call_member(buf(), "isEmpty", vec![])),
        getter("isNotEmpty", "bool", call_member(buf(), "isNotEmpty", vec![])),
        method(
            "write",
            vec![param("o", Some("Object"), None)],
            Some("void"),
            vec![set_buf(concat(buf(), stringify(ident("o"))))],
        ),
        method(
            "writeln",
            vec![param("o", Some("Object"), Some(str_lit("")))],
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
            vec![Statement::with_span(
                StmtKind::Return(Some(buf())),
                span(),
            )],
        ),
    ];
    Statement::with_span(
        StmtKind::ClassDecl {
            name: "StringBuffer".to_string(),
            parents: Vec::new(),
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
            decorators: vec![] },
        span(),
    )
}
