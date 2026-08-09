//! Dart walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//! Once this returns a `Module`, the rest of the compilation pipeline
//! (compile_class / compile_expression / etc.) is shared with every
//! other vybex language and works without any Dart-specific knowledge.
//!
//! ## Notes on Dart semantics that the walker normalises
//!
//! - **`this.field` constructor params**: When a constructor has `this.x`,
//!   `this.y` params, we synthesise assignments `this.x = x; this.y = y;`
//!   at the start of the constructor body. The `this.` prefix is stripped
//!   from param names.
//!
//! - **Constructor initializer lists** (`: super(args), field = expr`):
//!   `super(args)` is walked as base_args. `field = expr` assignments are
//!   prepended to the constructor body.
//!
//! - **Factory constructors**: Treated as static methods returning an instance.
//!
//! - **Cascade operator** (`..`): Desugared into a sequence of statements
//!   on the same object using a temp variable pattern.
//!
//! - **Named parameters**: Set `Argument { name: Some(label), value }`.
//!
//! - **`final`/`const` declarations**: Map to `VarDeclKind::Const` (immutable).
//!
//! - **`var`/typed declarations**: Map to `VarDeclKind::Let`.
//!
//! - **Enum declarations**: Each enum value becomes a class constant. Mapped
//!   to `StmtKind::ClassDecl` with `ClassMember::Const` entries.
//!
//! - **For-in**: Always `of: true` — Dart iterates values.
//!
//! - **Switch default**: Emitted as `SwitchCase { conditions: vec![] }` in
//!   source order (not separate `default` field).
//!
//! - **Mixins** (`with Mixin`): Treated as additional parent classes.
//!   Appended to `parents` list after the `extends` parent.

use super::{DartParser, Rule};
use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};
use vybe_ast::*;
use vybe_compiler::primitives::generics as common_generics;

const DART_USER_ADD_METHOD: &str = "__dart_user_add";

thread_local! {
    /// Types the program itself declares (class/enum/mixin/extension). Flutter
    /// named-constructor desugaring skips any type present here — the user's
    /// own declaration wins over the built-in allowlist (shadowing).
    static USER_DECLARED_TYPES: std::cell::RefCell<HashSet<String>> =
        std::cell::RefCell::new(HashSet::new());

    /// Names declared with `mixin`. Read by `normalize_class` to declare
    /// `Augmentation` records for `class X with M` — the shared model that
    /// replaces per-language folding (flexclassplan.md §4c).
    static DART_MIXIN_NAMES: std::cell::RefCell<HashSet<String>> =
        std::cell::RefCell::new(HashSet::new());

    /// `class name -> mixins it declared`, recorded by `apply_mixins` as it
    /// strips them from the parent list. `normalize_class` reads this to
    /// declare `Augmentation` records: by the time the compiler sees the
    /// ClassDecl the mixins are gone from `parents`, since a mixin is not a
    /// superclass.
    static DART_CLASS_MIXINS: std::cell::RefCell<HashMap<String, Vec<String>>> =
        std::cell::RefCell::new(HashMap::new());

    /// The subset of [`USER_DECLARED_TYPES`] declared with `class` — i.e. the
    /// names that `Name(args)` CONSTRUCTS. Dart has no `new` keyword, so
    /// construction parses as an ordinary `Call`; this drives the rewrite to
    /// the common AST's `ExprKind::New` (see [`dart_call_or_new`]). Kept
    /// separate from the full type set because `enum`/`mixin`/`extension`
    /// declarations are not constructible.
    static USER_DECLARED_CLASSES: std::cell::RefCell<HashSet<String>> =
        std::cell::RefCell::new(HashSet::new());
}

/// Cheap source pre-scan for declared type names, so [`dart_flutter_named_ctor`]
/// can respect user shadowing. A declaration is a line starting with
/// `class`/`enum`/`mixin`/`extension` (with the usual `abstract`/`sealed`/… )
/// followed by the name.
/// Returns `(all declared types, the `class`-declared subset)`.
fn collect_user_declared_types(source: &str) -> (HashSet<String>, HashSet<String>) {
    let mut set = HashSet::new();
    let mut classes = HashSet::new();
    for line in source.lines() {
        let mut t = line.trim_start();
        for modifier in ["abstract ", "sealed ", "final ", "base ", "interface "] {
            if let Some(rest) = t.strip_prefix(modifier) {
                t = rest.trim_start();
            }
        }
        for kw in ["class ", "enum ", "mixin ", "extension "] {
            if let Some(rest) = t.strip_prefix(kw) {
                let name: String = rest
                    .trim_start()
                    .chars()
                    .take_while(|c| c.is_alphanumeric() || *c == '_')
                    .collect();
                if !name.is_empty() {
                    if kw == "class " {
                        classes.insert(name.clone());
                    }
                    set.insert(name);
                }
            }
        }
    }
    (set, classes)
}

/// True when `name` is a class the program declares — the names `Name(args)`
/// constructs.
/// True when `name` was declared with `mixin`.
pub(crate) fn is_dart_mixin(name: &str) -> bool {
    DART_MIXIN_NAMES.with(|m| m.borrow().contains(name))
}

/// The mixins `class_name` declared with `with`, in source order.
pub(crate) fn dart_class_mixins(class_name: &str) -> Vec<String> {
    DART_CLASS_MIXINS.with(|m| m.borrow().get(class_name).cloned().unwrap_or_default())
}

fn is_user_declared_class(name: &str) -> bool {
    USER_DECLARED_CLASSES.with(|s| s.borrow().contains(name))
}

/// Build a call expression, normalising `ClassName(args)` — a call whose callee
/// names a class the program declares — to the common AST's `ExprKind::New`.
///
/// Dart has no `new` keyword, so construction parses as an ordinary `Call`.
/// JS/PHP emit `New` from their `new` syntax and Python's walker performs this
/// same rewrite (`call_or_new`), so every shared consumer that keys on `New` —
/// `infer_expr_type_hint`, and through it receiver typing and user-method
/// resolution — saw construction in every language except Dart. That is why a
/// user method colliding with a profile value-method (`length`, `keys`, `last`,
/// …) lost to the value-method table on a freshly-constructed receiver
/// (`K().length()`) but won through a variable (`var k = K(); k.length()`).
///
/// Any other callee stays a plain `Call`.
fn dart_call_or_new(callee: Expression, args: Vec<Argument>) -> ExprKind {
    if let ExprKind::Ident(name) = &callee.kind {
        // A `dart:core` TYPE constructs like any class. The shared ctor path
        // (`ExprKind::New` → `lookup_type_ctor_target`) is what emits the
        // backing call AND stamps `__type`/`__types`; a plain `Call` reaches
        // neither, which is what left every builtin class anonymous. A user
        // declaration of the same name still wins — it is checked first.
        if is_user_declared_class(name) || crate::core_classes::is_core_class(name) {
            return ExprKind::New {
                class: Box::new(callee),
                args,
            };
        }
    }
    ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    }
}

/// Desugar an allowlisted Flutter named constructor (`Type.name(args)`) into
/// the primary catalog construction (`Type(field: value, …)`), which the
/// common resolver's `Ctor` path lowers. `Type.name(...)` is syntactically a
/// static-method call, so this is a strictly CLOSED allowlist — never a
/// blanket rewrite — and it bails when the user declares their own `Type`.
/// Returns the rewritten `Call`, or `None` when nothing matches.
/// True when `name` is a widget/value type registered in the Flutter catalog.
fn is_flutter_catalog_class(name: &str) -> bool {
    vybe_platform_flutter::emitter::flutter_classes()
        .iter()
        .any(|c| c.name == name)
}

/// `Rect` geometry methods, lowered to arithmetic over the receiver's edges.
///
/// Only names that cannot belong to another Dart type are handled: `inflate`,
/// `deflate`, `expandToInclude`, `intersect` and `overlaps` are `Rect`-only,
/// whereas `contains`/`isEmpty`/`center` also exist on `List`/`String`/`Map`
/// and would be mis-rewritten without a receiver type (which this pass does not
/// have). The receiver must be a plain identifier, since it is referenced
/// several times and must not be re-evaluated.
fn dart_rect_method(receiver: &Expression, method: &str, args: &[Argument]) -> Option<ExprKind> {
    if !matches!(receiver.kind, ExprKind::Ident(_)) {
        return None;
    }
    // ONLY names that no other Dart type defines. `contains`/`isEmpty`/`center`
    // are deliberately absent: they are equally `List`/`String`/`Map` members,
    // and deciding by name alone would miscompile ordinary `list.contains(x)`.
    // Handling them needs the receiver's DECLARED type, which belongs in the
    // scope the compiler already maintains (`Scope`/`Local.type_hint`), not in
    // a name table invented here.
    if !matches!(
        method,
        "inflate" | "deflate" | "expandToInclude" | "intersect" | "overlaps"
    ) {
        return None;
    }
    let edge = |e: &Expression, f: &str| {
        Expression::new(ExprKind::Member {
            object: Box::new(e.clone()),
            field: f.to_string(),
            null_safe: false,
        })
    };
    let bin = |op: BinOp, l: Expression, r: Expression| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(l),
            right: Box::new(r),
        })
    };
    // `a < b ? a : b` / `a > b ? a : b`
    let pick = |a: Expression, b: Expression, smaller: bool| {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(bin(
                if smaller { BinOp::Lt } else { BinOp::Gt },
                a.clone(),
                b.clone(),
            )),
            then: Box::new(a),
            else_: Box::new(b),
        })
    };
    let rect_from = |l: Expression, t: Expression, r: Expression, b: Expression| ExprKind::Call {
        callee: Box::new(Expression::ident("Rect")),
        args: rect_fields(l, t, r, b),
        optional: false,
    };
    let (l, t, r, b) = (
        edge(receiver, "left"),
        edge(receiver, "top"),
        edge(receiver, "right"),
        edge(receiver, "bottom"),
    );

    match method {
        // Grow (or shrink) every edge by the same delta.
        "inflate" | "deflate" => {
            let d = args.first()?.value.clone();
            let (out, inn) = if method == "inflate" {
                (BinOp::Sub, BinOp::Add)
            } else {
                (BinOp::Add, BinOp::Sub)
            };
            Some(rect_from(
                bin(out, l, d.clone()),
                bin(out, t, d.clone()),
                bin(inn, r, d.clone()),
                bin(inn, b, d),
            ))
        }
        // The smallest rect covering both.
        "expandToInclude" => {
            let o = args.first()?.value.clone();
            Some(rect_from(
                pick(l, edge(&o, "left"), true),
                pick(t, edge(&o, "top"), true),
                pick(r, edge(&o, "right"), false),
                pick(b, edge(&o, "bottom"), false),
            ))
        }
        // The overlapping region (empty/inverted when they do not overlap).
        "intersect" => {
            let o = args.first()?.value.clone();
            Some(rect_from(
                pick(l, edge(&o, "left"), false),
                pick(t, edge(&o, "top"), false),
                pick(r, edge(&o, "right"), true),
                pick(b, edge(&o, "bottom"), true),
            ))
        }
        // True when the two rects share any area.
        "overlaps" => {
            let o = args.first()?.value.clone();
            let and = |x: Expression, y: Expression| bin(BinOp::And, x, y);
            Some(
                and(
                    and(
                        bin(BinOp::Lt, l, edge(&o, "right")),
                        bin(BinOp::Gt, r, edge(&o, "left")),
                    ),
                    and(
                        bin(BinOp::Lt, t, edge(&o, "bottom")),
                        bin(BinOp::Gt, b, edge(&o, "top")),
                    ),
                )
                .kind,
            )
        }
        // Half-open containment: the right/bottom edges are exclusive.
        "contains" => {
            let p = args.first()?.value.clone();
            let (px, py) = (edge(&p, "dx"), edge(&p, "dy"));
            let and = |x: Expression, y: Expression| bin(BinOp::And, x, y);
            Some(
                and(
                    and(bin(BinOp::GtEq, px.clone(), l), bin(BinOp::Lt, px, r)),
                    and(bin(BinOp::GtEq, py.clone(), t), bin(BinOp::Lt, py, b)),
                )
                .kind,
            )
        }
        // Getters, so they arrive with no argument list.
        "center" => Some(ExprKind::Call {
            callee: Box::new(Expression::ident("Offset")),
            args: vec![
                Argument::positional(bin(
                    BinOp::Div,
                    bin(BinOp::Add, l, r),
                    Expression::new(ExprKind::Lit(Literal::Float(2.0))),
                )),
                Argument::positional(bin(
                    BinOp::Div,
                    bin(BinOp::Add, t, b),
                    Expression::new(ExprKind::Lit(Literal::Float(2.0))),
                )),
            ],
            optional: false,
        }),
        "isEmpty" => Some(bin(BinOp::Or, bin(BinOp::GtEq, l, r), bin(BinOp::GtEq, t, b)).kind),
        _ => None,
    }
}

/// `Color(packed).alpha/red/green/blue` — the four channels sliced out of the
/// packed ARGB word. Flutter computes them in getters; a `Color` is immutable,
/// so they are derived once at construction.
fn color_channel_args(packed: Expression) -> Vec<Argument> {
    let named = |field: &str, value: Expression| Argument {
        value,
        name: Some(field.to_string()),
        by_ref: false,
        spread: false,
    };
    // `(packed ~/ place) % 256`. Deliberately arithmetic, NOT `>>`/`&`: a
    // fully-opaque colour (`0xFF112233`) exceeds 2^31, so the bitwise operators
    // would wrap it negative and every channel would come out wrong.
    let channel = |place: i64| -> Expression {
        let divided = if place == 1 {
            packed.clone()
        } else {
            Expression::new(ExprKind::Binary {
                op: BinOp::IDiv,
                left: Box::new(packed.clone()),
                right: Box::new(Expression::int(place)),
            })
        };
        Expression::new(ExprKind::Binary {
            op: BinOp::Mod,
            left: Box::new(divided),
            right: Box::new(Expression::int(256)),
        })
    };
    vec![
        Argument::positional(packed.clone()),
        named("alpha", channel(0x1000000)),
        named("red", channel(0x10000)),
        named("green", channel(0x100)),
        named("blue", channel(1)),
    ]
}

/// Pack `a, r, g, b` (each 0-255) into the single ARGB word a `Color` stores.
fn pack_argb(a: Expression, r: Expression, g: Expression, b: Expression) -> Expression {
    // Multiply-and-add rather than shift-or: `alpha << 24` overflows a signed
    // 32-bit word for any opaque colour, which would store a negative `value`.
    let scale = |v: Expression, place: i64| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Mul,
            left: Box::new(v),
            right: Box::new(Expression::int(place)),
        })
    };
    let add = |l: Expression, r: Expression| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(l),
            right: Box::new(r),
        })
    };
    add(
        add(scale(a, 0x1000000), scale(r, 0x10000)),
        add(scale(g, 0x100), b),
    )
}

/// `dart:typed_data` list → the ECMA typed array that backs it. Dart's typed
/// lists ARE the ECMA typed arrays (`Uint8List(3)` is a 3-byte buffer exactly
/// like `new Uint8Array(3)`), and the shared compiler already constructs the
/// ECMA names, so the frontend just renames them.
fn dart_typed_list_alias(name: &str) -> Option<&'static str> {
    Some(match name {
        "Uint8List" => "Uint8Array",
        "Uint8ClampedList" => "Uint8ClampedArray",
        "Int8List" => "Int8Array",
        "Uint16List" => "Uint16Array",
        "Int16List" => "Int16Array",
        "Uint32List" => "Uint32Array",
        "Int32List" => "Int32Array",
        "Float32List" => "Float32Array",
        "Float64List" => "Float64Array",
        "BigInt64List" => "BigInt64Array",
        "BigUint64List" => "BigUint64Array",
        _ => return None,
    })
}

/// `left op right` as an arithmetic expression — used to derive a `Rect`'s
/// edges and extents from the form its constructor was written in.
fn sub_or_add(left: Expression, right: Expression, op: BinOp) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

/// The six `Rect` fields for a set of edges: the four edges as given, plus the
/// derived `width`/`height`. Flutter exposes those as getters; a `Rect` is
/// immutable, so deriving them once at construction is equivalent.
fn rect_fields(
    left: Expression,
    top: Expression,
    right: Expression,
    bottom: Expression,
) -> Vec<Argument> {
    let arg = |field: &str, value: Expression| Argument {
        value,
        name: Some(field.to_string()),
        by_ref: false,
        spread: false,
    };
    vec![
        arg("left", left.clone()),
        arg("top", top.clone()),
        arg("right", right.clone()),
        arg("bottom", bottom.clone()),
        arg("width", sub_or_add(right, left, BinOp::Sub)),
        arg("height", sub_or_add(bottom, top, BinOp::Sub)),
    ]
}

/// Parse a catalog default-value source into an expression. Defaults are
/// written as ordinary Dart source in the widget modules, but they are a small
/// closed set (a number, a bool, an empty list, an enum constant), so they are
/// read directly rather than run back through the parser.
fn dart_default_expr(src: &str) -> Option<Expression> {
    let src = src.trim();
    Some(match src {
        "true" => Expression::new(ExprKind::Lit(Literal::Bool(true))),
        "false" => Expression::new(ExprKind::Lit(Literal::Bool(false))),
        "null" => Expression::null(),
        "const []" | "[]" => Expression::new(ExprKind::Array(Vec::new())),
        _ => {
            if let Ok(i) = src.parse::<i64>() {
                Expression::new(ExprKind::Lit(Literal::Int(i)))
            } else if let Ok(f) = src.parse::<f64>() {
                Expression::new(ExprKind::Lit(Literal::Float(f)))
            } else if let Some((enum_name, value)) = src.split_once('.') {
                // An enum-constant default (`FlexFit.loose`) folds exactly as a
                // written one does, so a default and a supplied value compare
                // equal.
                Expression::new(dart_flutter_enum_constant(enum_name, value)?)
            } else {
                return None;
            }
        }
    })
}

/// Apply a Flutter widget's constructor defaults: append a named argument for
/// every catalog field that declares a default and was not supplied. Flutter
/// itself defaults these (`Flexible().flex == 1`,
/// `Flexible().fit == FlexFit.loose`), and the shared construction path stores
/// whatever it is handed — so the frontend fills them in, which is where Dart's
/// own default-argument semantics belong.
fn inject_flutter_defaults(class_name: &str, args: &mut Vec<Argument>) {
    if USER_DECLARED_TYPES.with(|s| s.borrow().contains(class_name)) {
        return;
    }
    let supplied_positionally = args.iter().filter(|a| a.name.is_none()).count();
    for (field, default_src) in vybe_platform_flutter::emitter::field_defaults(class_name) {
        // Already given by name?
        if args.iter().any(|a| a.name.as_deref() == Some(field)) {
            continue;
        }
        // Or filled by a positional slot? Positional params come first in the
        // catalog's field order, so a field at index < positional-count is
        // already covered.
        let field_index = vybe_platform_flutter::emitter::flutter_classes()
            .iter()
            .find(|c| c.name == class_name)
            .and_then(|c| c.fields.iter().position(|f| f.name == field));
        if field_index.is_some_and(|i| i < supplied_positionally) {
            continue;
        }
        if let Some(value) = dart_default_expr(default_src) {
            args.push(Argument {
                value,
                name: Some(field.to_string()),
                by_ref: false,
                spread: false,
            });
        }
    }
}

/// Fold a Flutter enum constant `Enum.value` to the string Dart's `toString()`
/// produces for it (`"Clip.antiAlias"`). Enum constants are compile-time known,
/// so the canonical spelling IS the value: `==` between two constants is string
/// equality, a captured field prints as Dart prints it, and a catalog default
/// (`Column().direction == Axis.vertical`) compares equal. Returns `None` for
/// anything that is not a catalog enum, and for a user-declared shadow.
fn dart_flutter_enum_constant(enum_name: &str, value: &str) -> Option<ExprKind> {
    if USER_DECLARED_TYPES.with(|s| s.borrow().contains(enum_name)) {
        return None;
    }
    vybe_platform_flutter::emitter::enum_value_index(enum_name, value)?;
    Some(ExprKind::Lit(Literal::Str(format!("{enum_name}.{value}"))))
}

/// Fold `.name` / `.index` read off an already-folded enum constant. `expr` is
/// the folded `"Enum.value"` literal, so both answers are compile-time known.
fn dart_flutter_enum_member(expr: &Expression, field: &str) -> Option<ExprKind> {
    if field != "name" && field != "index" {
        return None;
    }
    let ExprKind::Lit(Literal::Str(s)) = &expr.kind else {
        return None;
    };
    let (enum_name, value) = s.split_once('.')?;
    let index = vybe_platform_flutter::emitter::enum_value_index(enum_name, value)?;
    Some(match field {
        "name" => ExprKind::Lit(Literal::Str(value.to_string())),
        _ => ExprKind::Lit(Literal::Int(index as i64)),
    })
}

fn dart_flutter_named_ctor(type_name: &str, ctor: &str, args: &[Argument]) -> Option<ExprKind> {
    if USER_DECLARED_TYPES.with(|s| s.borrow().contains(type_name)) {
        return None;
    }
    fn named(field: &str, value: Expression) -> Argument {
        Argument {
            value,
            name: Some(field.to_string()),
            by_ref: false,
            spread: false,
        }
    }
    fn zero() -> Expression {
        Expression::new(ExprKind::Lit(Literal::Float(0.0)))
    }
    let positional: Vec<Expression> = args
        .iter()
        .filter(|a| a.name.is_none())
        .map(|a| a.value.clone())
        .collect();
    let by_name = |label: &str| -> Option<Expression> {
        args.iter()
            .find(|a| a.name.as_deref() == Some(label))
            .map(|a| a.value.clone())
    };
    let construct = |ty: &str, fields: Vec<Argument>| -> ExprKind {
        ExprKind::Call {
            callee: Box::new(Expression::ident(ty)),
            args: fields,
            optional: false,
        }
    };

    match (type_name, ctor) {
        // EdgeInsets: four resolved edges. `.all`/`.symmetric`/`.only`/`.fromLTRB`.
        ("EdgeInsets" | "EdgeInsetsDirectional", "all") => {
            let v = positional.first()?.clone();
            Some(construct(
                "EdgeInsets",
                vec![
                    named("left", v.clone()),
                    named("top", v.clone()),
                    named("right", v.clone()),
                    named("bottom", v),
                ],
            ))
        }
        ("EdgeInsets" | "EdgeInsetsDirectional", "symmetric") => {
            let h = by_name("horizontal").unwrap_or_else(zero);
            let v = by_name("vertical").unwrap_or_else(zero);
            Some(construct(
                "EdgeInsets",
                vec![
                    named("left", h.clone()),
                    named("right", h),
                    named("top", v.clone()),
                    named("bottom", v),
                ],
            ))
        }
        ("EdgeInsets", "only") => Some(construct(
            "EdgeInsets",
            vec![
                named("left", by_name("left").unwrap_or_else(zero)),
                named("top", by_name("top").unwrap_or_else(zero)),
                named("right", by_name("right").unwrap_or_else(zero)),
                named("bottom", by_name("bottom").unwrap_or_else(zero)),
            ],
        )),
        ("EdgeInsets", "fromLTRB") => Some(construct(
            "EdgeInsets",
            vec![
                named("left", positional.first()?.clone()),
                named("top", positional.get(1)?.clone()),
                named("right", positional.get(2)?.clone()),
                named("bottom", positional.get(3)?.clone()),
            ],
        )),

        // Image factory ctors → `Image(image: <Provider>(src), …passthrough)`.
        ("Image", "network") | ("Image", "asset") | ("Image", "memory") | ("Image", "file") => {
            let provider = match ctor {
                "network" => "NetworkImage",
                "asset" => "AssetImage",
                "memory" => "MemoryImage",
                _ => "FileImage",
            };
            let src = positional.first()?.clone();
            let image = Expression::new(construct(provider, vec![Argument::positional(src)]));
            let mut fields = vec![named("image", image)];
            fields.extend(args.iter().filter(|a| a.name.is_some()).cloned());
            Some(construct("Image", fields))
        }

        // SizedBox factory ctors → explicit width/height (+ optional child).
        ("SizedBox", "shrink") => {
            let mut fields = vec![named("width", zero()), named("height", zero())];
            if let Some(c) = by_name("child") {
                fields.push(named("child", c));
            }
            Some(construct("SizedBox", fields))
        }
        ("SizedBox", "expand") => {
            let inf = || {
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident("double")),
                    field: "infinity".to_string(),
                    null_safe: false,
                })
            };
            let mut fields = vec![named("width", inf()), named("height", inf())];
            if let Some(c) = by_name("child") {
                fields.push(named("child", c));
            }
            Some(construct("SizedBox", fields))
        }
        ("SizedBox", "square") => {
            let d = by_name("dimension").unwrap_or_else(zero);
            let mut fields = vec![named("width", d.clone()), named("height", d)];
            if let Some(c) = by_name("child") {
                fields.push(named("child", c));
            }
            Some(construct("SizedBox", fields))
        }
        ("SizedBox", "fromSize") => {
            let size = by_name("size")?;
            let w = Expression::new(ExprKind::Member {
                object: Box::new(size.clone()),
                field: "width".to_string(),
                null_safe: false,
            });
            let h = Expression::new(ExprKind::Member {
                object: Box::new(size),
                field: "height".to_string(),
                null_safe: false,
            });
            Some(construct(
                "SizedBox",
                vec![named("width", w), named("height", h)],
            ))
        }

        // `Widget.canUpdate(old, new)` — a STATIC, not a constructor: two
        // widgets are interchangeable when they have the same runtime type and
        // the same key.
        ("Widget", "canUpdate") => {
            let (a, b) = (positional.first()?.clone(), positional.get(1)?.clone());
            let field = |o: &Expression, f: &str| {
                Expression::new(ExprKind::Member {
                    object: Box::new(o.clone()),
                    field: f.to_string(),
                    null_safe: false,
                })
            };
            let eq = |l: Expression, r: Expression| {
                Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(l),
                    right: Box::new(r),
                })
            };
            Some(
                Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(eq(field(&a, "__type"), field(&b, "__type"))),
                    right: Box::new(eq(field(&a, "key"), field(&b, "key"))),
                })
                .kind,
            )
        }

        // `Text.rich(span)` builds a Text whose content is a TextSpan tree
        // rather than a plain string.
        ("Text", "rich") => {
            let span = positional.first()?.clone();
            let mut fields = vec![named("textSpan", span)];
            fields.extend(args.iter().filter(|a| a.name.is_some()).cloned());
            Some(construct("Text", fields))
        }

        // `Color.fromARGB(a, r, g, b)` / `Color.fromRGBO(r, g, b, opacity)`
        // both pack into the same ARGB word the default constructor takes.
        ("Color", "fromARGB") => {
            let packed = pack_argb(
                positional.first()?.clone(),
                positional.get(1)?.clone(),
                positional.get(2)?.clone(),
                positional.get(3)?.clone(),
            );
            Some(construct("Color", color_channel_args(packed)))
        }
        ("Color", "fromRGBO") => {
            // Opacity is 0.0-1.0; the stored alpha is 0-255, rounded.
            let opacity = positional.get(3)?.clone();
            let alpha = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::Binary {
                        op: BinOp::Mul,
                        left: Box::new(opacity),
                        right: Box::new(Expression::new(ExprKind::Lit(Literal::Float(255.0)))),
                    })),
                    field: "round".to_string(),
                    null_safe: false,
                })),
                args: Vec::new(),
                optional: false,
            });
            let packed = pack_argb(
                alpha,
                positional.first()?.clone(),
                positional.get(1)?.clone(),
                positional.get(2)?.clone(),
            );
            Some(construct("Color", color_channel_args(packed)))
        }

        // ── dart:ui geometry ────────────────────────────────────────────
        // A Rect stores four edges; `width`/`height` are derived once here so
        // they read back without a computed getter.
        ("Rect", "fromLTRB") => {
            let (l, t, r, b) = (
                positional.first()?.clone(),
                positional.get(1)?.clone(),
                positional.get(2)?.clone(),
                positional.get(3)?.clone(),
            );
            Some(construct("Rect", rect_fields(l, t, r, b)))
        }
        ("Rect", "fromLTWH") => {
            let (l, t) = (positional.first()?.clone(), positional.get(1)?.clone());
            let (w, h) = (positional.get(2)?.clone(), positional.get(3)?.clone());
            let right = sub_or_add(l.clone(), w, BinOp::Add);
            let bottom = sub_or_add(t.clone(), h, BinOp::Add);
            Some(construct("Rect", rect_fields(l, t, right, bottom)))
        }
        ("Rect", "fromPoints") => {
            let (a, b) = (positional.first()?.clone(), positional.get(1)?.clone());
            let m = |o: &Expression, f: &str| {
                Expression::new(ExprKind::Member {
                    object: Box::new(o.clone()),
                    field: f.to_string(),
                    null_safe: false,
                })
            };
            // The two corners may be given in any order, so the rect is
            // normalised: left/top take the smaller coordinate.
            let pick = |x: Expression, y: Expression, smaller: bool| {
                Expression::new(ExprKind::Ternary {
                    cond: Box::new(Expression::new(ExprKind::Binary {
                        op: if smaller { BinOp::Lt } else { BinOp::Gt },
                        left: Box::new(x.clone()),
                        right: Box::new(y.clone()),
                    })),
                    then: Box::new(x),
                    else_: Box::new(y),
                })
            };
            Some(construct(
                "Rect",
                rect_fields(
                    pick(m(&a, "dx"), m(&b, "dx"), true),
                    pick(m(&a, "dy"), m(&b, "dy"), true),
                    pick(m(&a, "dx"), m(&b, "dx"), false),
                    pick(m(&a, "dy"), m(&b, "dy"), false),
                ),
            ))
        }
        ("Rect", "fromCircle") => {
            let center = by_name("center")?;
            let radius = by_name("radius")?;
            let m = |f: &str| {
                Expression::new(ExprKind::Member {
                    object: Box::new(center.clone()),
                    field: f.to_string(),
                    null_safe: false,
                })
            };
            Some(construct(
                "Rect",
                rect_fields(
                    sub_or_add(m("dx"), radius.clone(), BinOp::Sub),
                    sub_or_add(m("dy"), radius.clone(), BinOp::Sub),
                    sub_or_add(m("dx"), radius.clone(), BinOp::Add),
                    sub_or_add(m("dy"), radius, BinOp::Add),
                ),
            ))
        }
        // `Radius.circular(r)` — equal axes; `.elliptical(x, y)` — explicit.
        ("Radius", "circular") => {
            let r = positional.first()?.clone();
            Some(construct(
                "Radius",
                vec![named("x", r.clone()), named("y", r)],
            ))
        }
        ("Radius", "elliptical") => Some(construct(
            "Radius",
            vec![
                named("x", positional.first()?.clone()),
                named("y", positional.get(1)?.clone()),
            ],
        )),
        // An RRect carries the rect's box plus a radius per corner.
        ("RRect", "fromRectAndRadius") | ("RRect", "fromRectAndCorners") => {
            let rect = positional.first().cloned().or_else(|| by_name("rect"))?;
            let m = |f: &str| {
                Expression::new(ExprKind::Member {
                    object: Box::new(rect.clone()),
                    field: f.to_string(),
                    null_safe: false,
                })
            };
            let uniform = positional.get(1).cloned().or_else(|| by_name("radius"));
            let corner = |name: &str| -> Expression {
                by_name(name)
                    .or_else(|| uniform.clone())
                    .unwrap_or_else(|| Expression::null())
            };
            let mut fields = rect_fields(m("left"), m("top"), m("right"), m("bottom"));
            // Flutter exposes each corner both as a `Radius` and as its two
            // scalar axes (`tlRadiusX`/`tlRadiusY`), so store all three.
            for (field, arg) in [
                ("tl", "topLeft"),
                ("tr", "topRight"),
                ("bl", "bottomLeft"),
                ("br", "bottomRight"),
            ] {
                let radius = corner(arg);
                let axis = |a: &str| {
                    Expression::new(ExprKind::Member {
                        object: Box::new(radius.clone()),
                        field: a.to_string(),
                        null_safe: false,
                    })
                };
                fields.push(named(&format!("{field}RadiusX"), axis("x")));
                fields.push(named(&format!("{field}RadiusY"), axis("y")));
                fields.push(named(&format!("{field}Radius"), radius));
            }
            Some(construct("RRect", fields))
        }
        ("RelativeRect", "fromLTRB") => Some(construct(
            "RelativeRect",
            vec![
                named("left", positional.first()?.clone()),
                named("top", positional.get(1)?.clone()),
                named("right", positional.get(2)?.clone()),
                named("bottom", positional.get(3)?.clone()),
            ],
        )),
        // BoxConstraints factories: `tight` pins both axes to a Size, `loose`
        // caps them, `expand` fills, `tightFor` pins only what is given.
        ("BoxConstraints", "tight") => {
            let size = positional.first().cloned().or_else(|| by_name("size"))?;
            let m = |f: &str| {
                Expression::new(ExprKind::Member {
                    object: Box::new(size.clone()),
                    field: f.to_string(),
                    null_safe: false,
                })
            };
            Some(construct(
                "BoxConstraints",
                vec![
                    named("minWidth", m("width")),
                    named("maxWidth", m("width")),
                    named("minHeight", m("height")),
                    named("maxHeight", m("height")),
                ],
            ))
        }
        ("BoxConstraints", "loose") => {
            let size = positional.first().cloned().or_else(|| by_name("size"))?;
            let m = |f: &str| {
                Expression::new(ExprKind::Member {
                    object: Box::new(size.clone()),
                    field: f.to_string(),
                    null_safe: false,
                })
            };
            Some(construct(
                "BoxConstraints",
                vec![
                    named("minWidth", zero()),
                    named("maxWidth", m("width")),
                    named("minHeight", zero()),
                    named("maxHeight", m("height")),
                ],
            ))
        }
        ("BoxConstraints", "tightFor") | ("BoxConstraints", "expand") => {
            let infinity = || {
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident("double")),
                    field: "infinity".to_string(),
                    null_safe: false,
                })
            };
            // `tightFor` leaves an unspecified axis unconstrained; `expand`
            // fills it.
            let unspecified = |e: Option<Expression>, fallback: Expression| e.unwrap_or(fallback);
            let w = by_name("width");
            let h = by_name("height");
            let (w_min, w_max) = match &w {
                Some(v) => (v.clone(), v.clone()),
                None if ctor == "expand" => (infinity(), infinity()),
                None => (zero(), infinity()),
            };
            let (h_min, h_max) = match &h {
                Some(v) => (v.clone(), v.clone()),
                None if ctor == "expand" => (infinity(), infinity()),
                None => (zero(), infinity()),
            };
            let _ = unspecified;
            Some(construct(
                "BoxConstraints",
                vec![
                    named("minWidth", w_min),
                    named("maxWidth", w_max),
                    named("minHeight", h_min),
                    named("maxHeight", h_max),
                ],
            ))
        }

        // Positioned.fill — stretch to fill the Stack: zero edges (any of
        // left/top/right/bottom may be overridden as a margin), plus child.
        ("Positioned", "fill") => {
            let mut fields = vec![
                named("left", by_name("left").unwrap_or_else(zero)),
                named("top", by_name("top").unwrap_or_else(zero)),
                named("right", by_name("right").unwrap_or_else(zero)),
                named("bottom", by_name("bottom").unwrap_or_else(zero)),
            ];
            if let Some(c) = by_name("child") {
                fields.push(named("child", c));
            }
            Some(construct("Positioned", fields))
        }
        // Positioned.fromRect — left/top + width/height from the rect.
        ("Positioned", "fromRect") => {
            let rect = by_name("rect")?;
            let m = |f: &str| {
                Expression::new(ExprKind::Member {
                    object: Box::new(rect.clone()),
                    field: f.to_string(),
                    null_safe: false,
                })
            };
            let mut fields = vec![
                named("left", m("left")),
                named("top", m("top")),
                named("width", m("width")),
                named("height", m("height")),
            ];
            if let Some(c) = by_name("child") {
                fields.push(named("child", c));
            }
            Some(construct("Positioned", fields))
        }
        // Positioned.fromRelativeRect — all four edges from the rect.
        ("Positioned", "fromRelativeRect") => {
            let rect = by_name("rect")?;
            let m = |f: &str| {
                Expression::new(ExprKind::Member {
                    object: Box::new(rect.clone()),
                    field: f.to_string(),
                    null_safe: false,
                })
            };
            let mut fields = vec![
                named("left", m("left")),
                named("top", m("top")),
                named("right", m("right")),
                named("bottom", m("bottom")),
            ];
            if let Some(c) = by_name("child") {
                fields.push(named("child", c));
            }
            Some(construct("Positioned", fields))
        }

        // General fallback: any other named constructor on a known Flutter
        // catalog class is an alternate entry point that captures the same
        // fields — forward the args to the base type's construction.
        // Any other named constructor on a catalog class is an alternate entry
        // point capturing the same fields — forward to the base construction.
        //
        // ABSTRACT bases are excluded: they have no constructor, so a call like
        // `Widget.canUpdate(a, b)` is a STATIC, and constructing a `Widget`
        // from it silently returned an object where a bool was expected.
        _ => {
            let class = vybe_platform_flutter::emitter::flutter_classes()
                .iter()
                .find(|c| c.name == type_name)?;
            let constructible = class.widget_host_fn.is_some() || !class.fields.is_empty();
            constructible.then(|| construct(type_name, args.to_vec()))
        }
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let source = normalize_dart_expression_source(source);
    // Record the program's own declared types up front so Flutter named-ctor
    // desugaring respects user shadowing.
    DART_CONST_POOL.with(|pool| pool.borrow_mut().clear());
    let (declared_types, declared_classes) = collect_user_declared_types(&source);
    USER_DECLARED_TYPES.with(|s| *s.borrow_mut() = declared_types);
    USER_DECLARED_CLASSES.with(|s| *s.borrow_mut() = declared_classes);
    // Flutter render runtime: the `platforms/flutter` adapter owns its Dart
    // runtime (`runApp` + the widget-tree realizer). Append it ONLY when a
    // program actually renders — imports a Flutter library AND references
    // `runApp`. Widget-only code (construction, `is`-checks, the TDD suite)
    // never imports it, so it stays out of those programs entirely.
    let source = if source.contains("package:flutter/") && source.contains("runApp") {
        format!("{source}\n{}", vybe_platform_flutter::runtime_source())
    } else {
        source
    };
    let mut pairs = DartParser::parse(Rule::program, &source)
        .map_err(|e| format!("Dart parse error: {}", e))?;
    let program = pairs.next().ok_or("empty parse")?;

    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut mixin_names: std::collections::HashSet<String> = std::collections::HashSet::new();

    for pair in program.into_inner() {
        match pair.as_rule() {
            Rule::EOI => continue,
            Rule::import_declaration => imports.push(walk_import(pair)?),
            _ => {
                let was_mixin = pair.as_rule() == Rule::mixin_declaration;
                if let Some(stmt) = walk_top_level(pair)? {
                    if was_mixin {
                        if let StmtKind::ClassDecl { ref name, .. } = stmt.kind {
                            mixin_names.insert(name.clone());
                        }
                    }
                    body.push(stmt);
                }
            }
        }
    }

    // Mixin merge: copy members from each mixin into classes that
    // declare `with Mixin` (parents). Walker normalisation so the
    // shared class compiler sees a single flat class instead of
    // multi-mixin inheritance.
    // `dart:core` classes the runtime provides, declared as ordinary AST so
    // they normalise and compile exactly like a user class — see
    // `core_classes/`. Prepended so they precede any `const` binding that
    // constructs one, and skipped for a name the program declares itself.
    //
    // Spliced HERE, before the class passes below, not after them. Those passes
    // build the static-type environment that types a member chain and an
    // operator result (`rewrite_user_add_methods` collects every class's method
    // return types), so a class inserted afterwards is invisible to them: its
    // `operator +` exists on the class but no expression using it was ever
    // typed, and `d1 + d2` reached the slot lookup with nothing to find.
    let core_classes =
        crate::core_classes::declarations_for(&source, |name| is_user_declared_class(name));
    if !core_classes.is_empty() {
        body.splice(0..0, core_classes);
    }

    DART_MIXIN_NAMES.with(|m| *m.borrow_mut() = mixin_names.clone());
    DART_CLASS_MIXINS.with(|m| m.borrow_mut().clear());
    apply_mixins(&mut body, &mixin_names);
    apply_inherited_concrete_members(&mut body, &mixin_names);
    rewrite_inherited_instance_member_idents(&mut body, &mixin_names);
    rewrite_user_add_methods(&mut body);
    // Route failed member access on a `dynamic` receiver to the object's
    // `noSuchMethod`. Runs last: it reads the finished class list to decide
    // whether the program uses the hook at all, and the mixin passes above can
    // still be what puts `noSuchMethod` on a class.
    apply_no_such_method(&mut body);

    // Const bindings go after the last top-level DECLARATION, not at the very
    // top: `const Token(1)` constructs a user class, and hoisting it above
    // `class Token` ran the constructor before the class existed.
    let const_decls = dart_const_pool_declarations();
    if !const_decls.is_empty() {
        let after_declarations = body
            .iter()
            .rposition(|stmt| {
                matches!(
                    stmt.kind,
                    StmtKind::ClassDecl { .. }
                        | StmtKind::EnumDecl { .. }
                        | StmtKind::InterfaceDecl { .. }
                        | StmtKind::StructDecl { .. }
                        | StmtKind::FunctionDecl { .. }
                )
            })
            .map_or(0, |idx| idx + 1);
        body.splice(after_declarations..after_declarations, const_decls);
    }

    Ok(Module {
        name: String::new(),
        language: Lang::Dart,
        body,
        imports,
        // `this` is ambient, as in ECMA — a Dart method never declares it.
        directives: vybe_ast::Directives {
            receiver_binding: Some(vybe_ast::ReceiverBinding::Ambient),
            ..Default::default()
        },
    })
}

fn normalize_dart_expression_source(source: &str) -> String {
    let source = normalize_runtime_type_generic_literals(source);
    normalize_parenthesized_is_ternary(&source)
}

fn normalize_parenthesized_is_ternary(source: &str) -> String {
    let mut out = String::with_capacity(source.len());
    let bytes = source.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] != b'(' {
            // Copy a whole UTF-8 CHARACTER. `bytes[i] as char` is a Latin-1
            // decode — it turned each byte of `é` into a separate char, so
            // EVERY non-ASCII Dart source reached the parser as mojibake:
            // `'café'` became `cafÃ©`, which is why `.length` was 5,
            // `codeUnitAt(3)` was 195 (the first UTF-8 byte) instead of 233,
            // and every runes/codeUnits test disagreed with Dart.
            let ch = source[i..]
                .chars()
                .next()
                .expect("index is a char boundary");
            out.push(ch);
            i += ch.len_utf8();
            continue;
        }
        let start = i;
        let mut j = i + 1;
        while j < bytes.len() && bytes[j].is_ascii_whitespace() {
            j += 1;
        }
        while j < bytes.len() && (bytes[j].is_ascii_alphanumeric() || bytes[j] == b'_') {
            j += 1;
        }
        let after_subject = j;
        while j < bytes.len() && bytes[j].is_ascii_whitespace() {
            j += 1;
        }
        // `source.get(..)`, NOT `&source[..]`: `j` indexes BYTES, and the scans
        // above stop at the first byte of whatever follows. When that is a
        // multi-byte character, `j + 2` lands INSIDE it and slicing panics —
        // `print('<emoji>')` was enough to kill the compiler. `get` returns
        // `None` on a non-boundary, which folds the old length check in too.
        if source.get(j..j + 2) != Some("is") {
            out.push('(');
            i = start + 1;
            continue;
        }
        let after_is = j + 2;
        if after_is < bytes.len()
            && (bytes[after_is].is_ascii_alphanumeric() || bytes[after_is] == b'_')
        {
            out.push('(');
            i = start + 1;
            continue;
        }
        j = after_is;
        while j < bytes.len() && bytes[j].is_ascii_whitespace() {
            j += 1;
        }
        while j < bytes.len() && (bytes[j].is_ascii_alphanumeric() || bytes[j] == b'_') {
            j += 1;
        }
        let after_type = j;
        while j < bytes.len() && bytes[j].is_ascii_whitespace() {
            j += 1;
        }
        if j < bytes.len() && bytes[j] == b'?' {
            out.push('(');
            out.push_str(&source[start..after_type]);
            out.push(')');
            out.push_str(&source[after_type..=j]);
            i = j + 1;
        } else {
            out.push('(');
            i = start + 1;
        }
        let _ = after_subject;
    }
    out
}

fn normalize_runtime_type_generic_literals(source: &str) -> String {
    let mut out = String::with_capacity(source.len());
    let bytes = source.as_bytes();
    let mut i = 0;
    while let Some(rel) = source[i..].find("runtimeType") {
        let start = i + rel;
        out.push_str(&source[i..start + "runtimeType".len()]);
        i = start + "runtimeType".len();

        let mut j = i;
        while j < bytes.len() && bytes[j].is_ascii_whitespace() {
            j += 1;
        }
        let is_cmp = j + 1 < bytes.len()
            && ((bytes[j] == b'=' && bytes[j + 1] == b'=')
                || (bytes[j] == b'!' && bytes[j + 1] == b'='));
        if !is_cmp {
            continue;
        }

        out.push_str(&source[i..j + 2]);
        j += 2;
        while j < bytes.len() && bytes[j].is_ascii_whitespace() {
            out.push(bytes[j] as char);
            j += 1;
        }
        let ident_start = j;
        while j < bytes.len() && (bytes[j].is_ascii_alphanumeric() || bytes[j] == b'_') {
            j += 1;
        }
        if ident_start == j || j >= bytes.len() || bytes[j] != b'<' {
            i = ident_start;
            continue;
        }
        out.push_str(&source[ident_start..j]);
        let mut depth = 0i32;
        while j < bytes.len() {
            match bytes[j] {
                b'<' => depth += 1,
                b'>' => {
                    depth -= 1;
                    if depth == 0 {
                        j += 1;
                        break;
                    }
                }
                _ => {}
            }
            j += 1;
        }
        i = j;
    }
    out.push_str(&source[i..]);
    out
}

/// Copy methods/fields from each `mixin Foo { ... }` into every
/// `class X with Foo, Bar` (or `class X extends Base with Foo`) and
/// strip the mixin names out of the class's parent list. Mixins
/// themselves stay in the body — they're harmless ClassDecls and
/// some user code may reference them by name.
fn apply_mixins(body: &mut Vec<Statement>, mixin_names: &std::collections::HashSet<String>) {
    if mixin_names.is_empty() {
        return;
    }
    let mut class_members_by_name: HashMap<String, Vec<ClassMember>> = HashMap::new();
    let mut class_parents_by_name: HashMap<String, Vec<String>> = HashMap::new();
    for stmt in body.iter() {
        if let StmtKind::ClassDecl {
            name,
            parents,
            members,
            ..
        } = &stmt.kind
        {
            class_members_by_name.insert(name.clone(), members.clone());
            class_parents_by_name.insert(name.clone(), parents.clone());
        }
    }

    // First pass: a bare identifier in a `mixin M on T` body that names a
    // member of `T` is `this.<member>` — the mixin declares no such member of
    // its own, so nothing else can resolve it.
    //
    // Rewritten IN PLACE. This used to rewrite a CLONE that only the walker's
    // member-copy consumed; once that copy moved to the shared augmentation
    // pass the clone became dead, and with it every `on`-supertype reference.
    // The mixin's own `ClassDecl` is what the shared pass reads now, so the
    // rewrite has to land there.
    let mut on_supertype_members: HashMap<String, Vec<String>> = HashMap::new();
    for stmt in body.iter() {
        if let StmtKind::ClassDecl { name, parents, .. } = &stmt.kind {
            if mixin_names.contains(name) {
                on_supertype_members.insert(
                    name.clone(),
                    collect_instance_member_names_for_types(
                        parents,
                        &class_members_by_name,
                        &class_parents_by_name,
                    ),
                );
            }
        }
    }
    let declared_mixins: std::collections::HashSet<String> =
        on_supertype_members.keys().cloned().collect();
    for stmt in body.iter_mut() {
        if let StmtKind::ClassDecl { name, members, .. } = &mut stmt.kind {
            if let Some(extra_members) = on_supertype_members.get(name) {
                let extra_refs: Vec<&str> = extra_members.iter().map(String::as_str).collect();
                rewrite_instance_member_idents(members, &extra_refs);
            }
        }
    }
    // Second pass: merge into consumers.
    for stmt in body.iter_mut() {
        if let StmtKind::ClassDecl {
            name: cname,
            parents,
            members,
            ..
        } = &mut stmt.kind
        {
            if mixin_names.contains(cname) {
                continue;
            }
            let original_member_names: std::collections::HashSet<String> =
                members.iter().filter_map(member_name).collect();
            let mut new_parents = Vec::new();
            for parent in parents.drain(..) {
                if declared_mixins.contains(&parent) {
                    DART_CLASS_MIXINS.with(|m| {
                        m.borrow_mut()
                            .entry(cname.clone())
                            .or_default()
                            .push(parent.clone())
                    });
                    // MEMBER COPYING REMOVED — the shared augmentation pass
                    // (`primitives/class_augmentation.rs`) folds these now, from
                    // the `Augmentation` records `dart/normalize_class.rs`
                    // declares. Only the parent-stripping below is still the
                    // walker's job: a mixin is not a superclass.
                    let _ = &original_member_names;
                } else {
                    new_parents.push(parent);
                }
            }
            *parents = new_parents;
        }
    }
}

fn same_name_property_member_exists(members: &[ClassMember], incoming: &ClassMember) -> bool {
    let ClassMember::Property { name, .. } = incoming else {
        return false;
    };
    members.iter().any(|existing| {
        matches!(existing, ClassMember::Property { name: existing_name, .. } if existing_name == name)
    })
}

fn collect_instance_member_names_for_types(
    type_names: &[String],
    class_members_by_name: &HashMap<String, Vec<ClassMember>>,
    class_parents_by_name: &HashMap<String, Vec<String>>,
) -> Vec<String> {
    let mut names = Vec::new();
    let mut seen_names = HashSet::new();
    let mut seen_types = HashSet::new();
    for type_name in type_names {
        collect_instance_member_names_for_type(
            type_name,
            class_members_by_name,
            class_parents_by_name,
            &mut seen_types,
            &mut seen_names,
            &mut names,
        );
    }
    names
}

fn collect_instance_member_names_for_type(
    type_name: &str,
    class_members_by_name: &HashMap<String, Vec<ClassMember>>,
    class_parents_by_name: &HashMap<String, Vec<String>>,
    seen_types: &mut HashSet<String>,
    seen_names: &mut HashSet<String>,
    names: &mut Vec<String>,
) {
    if !seen_types.insert(type_name.to_string()) {
        return;
    }
    if let Some(parents) = class_parents_by_name.get(type_name) {
        for parent in parents {
            collect_instance_member_names_for_type(
                parent,
                class_members_by_name,
                class_parents_by_name,
                seen_types,
                seen_names,
                names,
            );
        }
    }
    if let Some(members) = class_members_by_name.get(type_name) {
        for member in members {
            if let Some(name) = instance_member_name(member) {
                if seen_names.insert(name.clone()) {
                    names.push(name);
                }
            }
        }
    }
}

fn rewrite_user_add_methods(body: &mut Vec<Statement>) {
    let mut add_return_types: HashMap<String, Option<String>> = HashMap::new();
    let mut operator_return_types: HashMap<(String, String), Option<String>> = HashMap::new();
    // Seed Flutter widget/value-type field types (from the flutter catalog) so
    // `double` fields render `.0` and chained value reads resolve. Feeds the
    // SAME static-type tracker as operator overloading — display is driven by
    // declared types, never a runtime check. Harmless for non-Flutter programs
    // (those type names never appear); user classes are inserted below and win
    // on key collision (insert-after).
    for (owner, field, ty) in vybe_platform_flutter::emitter::field_type_seed() {
        operator_return_types.insert((owner.to_string(), field.to_string()), Some(ty.to_string()));
    }
    let mut iterator_return_classes: HashMap<String, String> = HashMap::new();
    let mut iterator_current_types: HashMap<String, String> = HashMap::new();
    let mut class_parents: Vec<(String, Vec<String>)> = Vec::new();
    for stmt in body.iter() {
        if let StmtKind::ClassDecl { name, members, .. } = &stmt.kind {
            if let StmtKind::ClassDecl { parents, .. } = &stmt.kind {
                class_parents.push((name.clone(), parents.clone()));
            }
            for member in members {
                if let ClassMember::Method(method) = member {
                    if let StmtKind::FunctionDecl {
                        name: method_name,
                        return_type,
                        modifiers,
                        ..
                    } = &method.kind
                    {
                        if !modifiers.is_static {
                            operator_return_types
                                .insert((name.clone(), method_name.clone()), return_type.clone());
                        }
                        if method_name == "add" && !modifiers.is_static {
                            add_return_types.insert(name.clone(), return_type.clone());
                            operator_return_types.insert(
                                (name.clone(), DART_USER_ADD_METHOD.to_string()),
                                return_type.clone(),
                            );
                        }
                    }
                } else if let ClassMember::Field {
                    name: field_name,
                    type_hint: Some(type_hint),
                    modifiers,
                    ..
                } = member
                {
                    if !modifiers.is_static {
                        operator_return_types
                            .insert((name.clone(), field_name.clone()), Some(type_hint.clone()));
                    }
                } else if let ClassMember::Property {
                    name: property_name,
                    type_hint,
                    getter: Some(_),
                    modifiers,
                    ..
                } = member
                {
                    if property_name == "current" && !modifiers.is_static {
                        if let Some(type_hint) = type_hint {
                            iterator_current_types.insert(name.clone(), type_hint.clone());
                        }
                    }
                    if !modifiers.is_static {
                        if let Some(type_hint) = type_hint {
                            operator_return_types.insert(
                                (name.clone(), property_name.clone()),
                                Some(type_hint.clone()),
                            );
                        }
                    }
                    if property_name == "hashCode" && !modifiers.is_static {
                        operator_return_types
                            .insert((name.clone(), "__get_hash".to_string()), type_hint.clone());
                    }
                    if property_name == "iterator" && !modifiers.is_static {
                        if let ClassMember::Property {
                            getter: Some(getter),
                            ..
                        } = member
                        {
                            if let Some(iterator_class) = dart_returned_class_from_body(getter) {
                                iterator_return_classes.insert(name.clone(), iterator_class);
                            }
                        }
                    }
                }
            }
        }
    }
    for (class_name, iterator_class) in &iterator_return_classes {
        if let Some(current_type) = iterator_current_types.get(iterator_class) {
            operator_return_types.insert(
                (class_name.clone(), "__dart_iter_element".to_string()),
                Some(current_type.clone()),
            );
        }
    }
    for _ in 0..class_parents.len() {
        let mut changed = false;
        for (class_name, parents) in &class_parents {
            if operator_return_types
                .contains_key(&(class_name.clone(), "__dart_iter_element".to_string()))
            {
                continue;
            }
            if let Some(parent_type) = parents.iter().find_map(|parent| {
                operator_return_types
                    .get(&(parent.clone(), "__dart_iter_element".to_string()))
                    .and_then(|ty| ty.clone())
            }) {
                operator_return_types.insert(
                    (class_name.clone(), "__dart_iter_element".to_string()),
                    Some(parent_type),
                );
                changed = true;
            }
        }
        if !changed {
            break;
        }
    }
    for stmt in body.iter_mut() {
        if let StmtKind::ClassDecl { members, .. } = &mut stmt.kind {
            for member in members {
                if let ClassMember::Method(method) = member {
                    if let StmtKind::FunctionDecl {
                        name, modifiers, ..
                    } = &mut method.kind
                    {
                        if name == "add" && !modifiers.is_static {
                            *name = DART_USER_ADD_METHOD.to_string();
                        }
                    }
                }
            }
        }
    }

    let mut env = HashMap::new();
    rewrite_user_add_calls_in_stmts(
        body,
        &mut env,
        None,
        &add_return_types,
        &operator_return_types,
    );
}

fn rewrite_user_add_calls_in_stmts(
    stmts: &mut [Statement],
    env: &mut HashMap<String, String>,
    current_class: Option<&str>,
    add_return_types: &HashMap<String, Option<String>>,
    operator_return_types: &HashMap<(String, String), Option<String>>,
) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::Expr(expr) => {
                rewrite_user_add_calls_in_expr(
                    expr,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
            StmtKind::Block(body) => {
                let mut block_env = env.clone();
                rewrite_user_add_calls_in_stmts(
                    body,
                    &mut block_env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        rewrite_user_add_calls_in_expr(
                            init,
                            env,
                            current_class,
                            add_return_types,
                            operator_return_types,
                        );
                    }
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if let Some(type_name) = decl
                            .type_hint
                            .as_deref()
                            .filter(|hint| {
                                dart_user_known_class(hint, add_return_types, operator_return_types)
                            })
                            .map(str::to_string)
                            .or_else(|| {
                                decl.init.as_ref().and_then(|init| {
                                    dart_user_add_expr_type(
                                        init,
                                        env,
                                        current_class,
                                        add_return_types,
                                        operator_return_types,
                                    )
                                })
                            })
                        {
                            env.insert(name.clone(), type_name);
                        }
                        // SplayTreeSet is a comparison-ordered set; track it so
                        // its `.add` routes to the shared sorted core.
                        let is_splay_tree_set = decl
                            .type_hint
                            .as_deref()
                            .map(|hint| hint.contains("SplayTreeSet"))
                            .unwrap_or(false)
                            || decl
                                .init
                                .as_ref()
                                .and_then(dart_constructor_expr_type)
                                .as_deref()
                                == Some("SplayTreeSet");
                        if is_splay_tree_set {
                            env.insert(name.clone(), "SplayTreeSet".to_string());
                        }
                    }
                }
            }
            StmtKind::FunctionDecl { body, .. } => {
                let mut fn_env = HashMap::new();
                rewrite_user_add_calls_in_stmts(
                    body,
                    &mut fn_env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
            StmtKind::ClassDecl { name, members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Method(method) => {
                            if let StmtKind::FunctionDecl { body, .. } = &mut method.kind {
                                let mut method_env = HashMap::new();
                                rewrite_user_add_calls_in_stmts(
                                    body,
                                    &mut method_env,
                                    Some(name),
                                    add_return_types,
                                    operator_return_types,
                                );
                            }
                        }
                        ClassMember::Constructor { body, .. } => {
                            let mut ctor_env = HashMap::new();
                            rewrite_user_add_calls_in_stmts(
                                body,
                                &mut ctor_env,
                                Some(name),
                                add_return_types,
                                operator_return_types,
                            );
                        }
                        ClassMember::Property { getter, setter, .. } => {
                            if let Some(body) = getter {
                                let mut prop_env = HashMap::new();
                                rewrite_user_add_calls_in_stmts(
                                    body,
                                    &mut prop_env,
                                    Some(name),
                                    add_return_types,
                                    operator_return_types,
                                );
                            }
                            if let Some(setter) = setter {
                                let mut prop_env = HashMap::new();
                                rewrite_user_add_calls_in_stmts(
                                    &mut setter.body,
                                    &mut prop_env,
                                    Some(name),
                                    add_return_types,
                                    operator_return_types,
                                );
                            }
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                rewrite_user_add_calls_in_expr(
                    cond,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
                let mut then_env = env.clone();
                rewrite_user_add_calls_in_stmts(
                    then_body,
                    &mut then_env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
                for (elif_cond, elif_body) in elifs {
                    rewrite_user_add_calls_in_expr(
                        elif_cond,
                        env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                    let mut elif_env = env.clone();
                    rewrite_user_add_calls_in_stmts(
                        elif_body,
                        &mut elif_env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
                if let Some(body) = else_body {
                    let mut else_env = env.clone();
                    rewrite_user_add_calls_in_stmts(
                        body,
                        &mut else_env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                let mut loop_env = env.clone();
                if let Some(init) = init {
                    rewrite_user_add_calls_in_stmts(
                        std::slice::from_mut(init),
                        &mut loop_env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
                if let Some(cond) = cond {
                    rewrite_user_add_calls_in_expr(
                        cond,
                        &mut loop_env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
                if let Some(update) = update {
                    rewrite_user_add_calls_in_expr(
                        update,
                        &mut loop_env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
                rewrite_user_add_calls_in_stmts(
                    body,
                    &mut loop_env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
            StmtKind::ForIn { iter, body, .. } => {
                rewrite_user_add_calls_in_expr(
                    iter,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
                let mut loop_env = env.clone();
                rewrite_user_add_calls_in_stmts(
                    body,
                    &mut loop_env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
            StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
                rewrite_user_add_calls_in_expr(
                    cond,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
                let mut loop_env = env.clone();
                rewrite_user_add_calls_in_stmts(
                    body,
                    &mut loop_env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
            StmtKind::Return(expr) | StmtKind::Throw { expr, .. } => {
                if let Some(expr) = expr {
                    rewrite_user_add_calls_in_expr(
                        expr,
                        env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
            }
            StmtKind::Assign { targets, value, .. } => {
                if targets.len() == 1 {
                    if let Some(call) = dart_user_index_set_call(
                        &targets[0],
                        value.clone(),
                        env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    ) {
                        stmt.kind = StmtKind::Expr(call);
                        continue;
                    }
                    if let Some(call) = dart_index_set_call(&targets[0], value.clone()) {
                        stmt.kind = StmtKind::Expr(call);
                        continue;
                    }
                }
                for target in targets {
                    rewrite_user_add_calls_in_expr(
                        target,
                        env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
                rewrite_user_add_calls_in_expr(
                    value,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                rewrite_user_add_calls_in_expr(
                    target,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
                rewrite_user_add_calls_in_expr(
                    value,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
            StmtKind::Switch {
                expr,
                cases,
                default,
            } => {
                rewrite_user_add_calls_in_expr(
                    expr,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
                for case in cases {
                    for cond in &mut case.conditions {
                        match cond {
                            CaseCondition::Value(value)
                            | CaseCondition::Comparison { expr: value, .. } => {
                                rewrite_user_add_calls_in_expr(
                                    value,
                                    env,
                                    current_class,
                                    add_return_types,
                                    operator_return_types,
                                );
                            }
                            CaseCondition::Range { from, to } => {
                                rewrite_user_add_calls_in_expr(
                                    from,
                                    env,
                                    current_class,
                                    add_return_types,
                                    operator_return_types,
                                );
                                rewrite_user_add_calls_in_expr(
                                    to,
                                    env,
                                    current_class,
                                    add_return_types,
                                    operator_return_types,
                                );
                            }
                        }
                    }
                    let mut case_env = env.clone();
                    rewrite_user_add_calls_in_stmts(
                        &mut case.body,
                        &mut case_env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
                if let Some(body) = default {
                    let mut default_env = env.clone();
                    rewrite_user_add_calls_in_stmts(
                        body,
                        &mut default_env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                let mut try_env = env.clone();
                rewrite_user_add_calls_in_stmts(
                    body,
                    &mut try_env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
                for catch in catches {
                    let mut catch_env = env.clone();
                    rewrite_user_add_calls_in_stmts(
                        &mut catch.body,
                        &mut catch_env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
                if let Some(body) = else_body {
                    let mut else_env = env.clone();
                    rewrite_user_add_calls_in_stmts(
                        body,
                        &mut else_env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
                if let Some(body) = finally {
                    let mut finally_env = env.clone();
                    rewrite_user_add_calls_in_stmts(
                        body,
                        &mut finally_env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
            }
            _ => {}
        }
    }
}

fn rewrite_user_add_calls_in_expr(
    expr: &mut Expression,
    env: &HashMap<String, String>,
    current_class: Option<&str>,
    add_return_types: &HashMap<String, Option<String>>,
    operator_return_types: &HashMap<(String, String), Option<String>>,
) {
    match &mut expr.kind {
        ExprKind::Member { object, field, .. } => {
            rewrite_user_add_calls_in_expr(
                object,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            if field == "add" {
                if let Some(type_name) = dart_user_add_expr_type(
                    object,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                ) {
                    if type_name == "SplayTreeSet" {
                        *field = "__dart_sorted_add".to_string();
                    } else if add_return_types.contains_key(&type_name) {
                        *field = DART_USER_ADD_METHOD.to_string();
                    }
                }
            } else if field == "hashCode" {
                let helper = if dart_user_add_expr_type(
                    object,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                )
                .is_some()
                {
                    "__dart_object_hash_code"
                } else {
                    "__dart_hash_code"
                };
                let object = (**object).clone();
                expr.kind = ExprKind::Call {
                    callee: Box::new(Expression::ident(helper)),
                    args: vec![Argument::positional(object)],
                    optional: false,
                };
            }
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_user_add_calls_in_expr(
                callee,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            for arg in &mut *args {
                rewrite_user_add_calls_in_expr(
                    &mut arg.value,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
            // `print(x)` where `x` is statically a `double` must render Dart
            // style (`10.0`, not `10`) — driven by the same static-type source
            // used for operator overloading, no runtime check.
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "print") {
                for arg in &mut *args {
                    if arg.name.is_none()
                        && matches!(
                            dart_user_add_expr_type(
                                &arg.value,
                                env,
                                current_class,
                                add_return_types,
                                operator_return_types,
                            )
                            .as_deref(),
                            Some("double")
                        )
                    {
                        let inner = arg.value.clone();
                        arg.value = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__dart_double_to_string")),
                            args: vec![Argument::positional(inner)],
                            optional: false,
                        });
                    }
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field == "toString" && args.is_empty() {
                    if matches!(
                        dart_user_add_expr_type(
                            object,
                            env,
                            current_class,
                            add_return_types,
                            operator_return_types,
                        )
                        .as_deref(),
                        Some("double")
                    ) {
                        expr.kind = ExprKind::Call {
                            callee: Box::new(Expression::ident("__dart_double_to_string")),
                            args: vec![Argument::positional((**object).clone())],
                            optional: false,
                        };
                        return;
                    }
                }
            }
            if let ExprKind::Member { object, field, .. } = &mut callee.kind {
                if field == "map" {
                    if let Some(type_name) = dart_user_add_expr_type(
                        object,
                        env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    ) {
                        if let Some(element_type) =
                            dart_iter_element_type(&type_name, operator_return_types)
                        {
                            if let Some(arg) = args.get_mut(0) {
                                rewrite_lambda_with_param_type(
                                    &mut arg.value,
                                    &element_type,
                                    env,
                                    current_class,
                                    add_return_types,
                                    operator_return_types,
                                );
                            }
                        }
                    }
                }
                if field == "expand" {
                    if let Some(type_name) = dart_user_add_expr_type(
                        object,
                        env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    ) {
                        if matches!(
                            dart_iter_element_type(&type_name, operator_return_types).as_deref(),
                            Some("List")
                        ) {
                            *field = "__dart_iter_expand_precurrent".to_string();
                        }
                    }
                }
            }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__dart_index_get")
                && args.len() == 2
            {
                let object = args[0].value.clone();
                let index = args[1].value.clone();
                if let Some(type_name) = dart_user_add_expr_type(
                    &object,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                ) {
                    if operator_return_types.contains_key(&(type_name, "__getitem__".to_string())) {
                        expr.kind = ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(object),
                                field: "__getitem__".to_string(),
                                null_safe: false,
                            })),
                            args: vec![Argument::positional(index)],
                            optional: false,
                        };
                    }
                }
            } else if matches!(
                &callee.kind,
                ExprKind::Member { object, field, .. }
                    if matches!(&object.kind, ExprKind::Ident(name) if name == "int")
                        && matches!(field.as_str(), "parse" | "tryParse")
            ) {
                for arg in args {
                    if matches!(arg.name.as_deref(), Some("radix")) {
                        arg.name = None;
                    }
                }
            } else if matches!(&callee.kind, ExprKind::Ident(name) if name == "__dart_eq")
                && args.len() == 2
            {
                let left = args[0].value.clone();
                let right = args[1].value.clone();
                if let Some(type_name) = dart_user_add_expr_type(
                    &left,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                ) {
                    if operator_return_types.contains_key(&(type_name, "__eq__".to_string())) {
                        expr.kind = ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(left),
                                field: "__eq__".to_string(),
                                null_safe: false,
                            })),
                            args: vec![Argument::positional(right)],
                            optional: false,
                        };
                    }
                }
            } else if !matches!(&callee.kind, ExprKind::Member { field, .. } if field == "call") {
                if let Some(type_name) = dart_user_add_expr_type(
                    callee,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                ) {
                    if operator_return_types.contains_key(&(type_name, "call".to_string())) {
                        let object = (**callee).clone();
                        let call_args = args.clone();
                        expr.kind = ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Index {
                                object: Box::new(object),
                                index: Box::new(Expression::string("call")),
                                null_safe: false,
                            })),
                            args: call_args,
                            optional: false,
                        };
                    }
                }
            }
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                rewrite_user_add_calls_in_expr(
                    &mut arg.value,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_user_add_calls_in_expr(
                left,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            rewrite_user_add_calls_in_expr(
                right,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            if let ExprKind::Binary { op, left, right } = &expr.kind {
                if let Some(method_name) = dart_user_binary_operator_method(op) {
                    if let Some(type_name) = dart_user_add_expr_type(
                        left,
                        env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    ) {
                        if operator_return_types.contains_key(&(type_name, method_name.to_string()))
                        {
                            expr.kind = ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Member {
                                    object: Box::new((**left).clone()),
                                    field: method_name.to_string(),
                                    null_safe: false,
                                })),
                                args: vec![Argument::positional((**right).clone())],
                                optional: false,
                            };
                        }
                    }
                }
            }
        }
        ExprKind::Unary { op, expr: inner } => {
            rewrite_user_add_calls_in_expr(
                inner,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            if let Some(method_name) = dart_user_unary_operator_method(op) {
                if let Some(type_name) = dart_user_add_expr_type(
                    inner,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                ) {
                    if operator_return_types.contains_key(&(type_name, method_name.to_string())) {
                        expr.kind = ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new((**inner).clone()),
                                field: method_name.to_string(),
                                null_safe: false,
                            })),
                            args: Vec::new(),
                            optional: false,
                        };
                    }
                }
            }
        }
        ExprKind::Await(inner)
        | ExprKind::Yield(Some(inner))
        | ExprKind::YieldFrom(inner)
        | ExprKind::Spread(inner)
        | ExprKind::RefLoad(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::Void(inner)
        | ExprKind::Delete(inner) => {
            rewrite_user_add_calls_in_expr(
                inner,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_user_add_calls_in_expr(
                cond,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            rewrite_user_add_calls_in_expr(
                then,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            rewrite_user_add_calls_in_expr(
                else_,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_user_add_calls_in_expr(
                object,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            rewrite_user_add_calls_in_expr(
                index,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            if let ExprKind::Index { object, index, .. } = &expr.kind {
                if let Some(type_name) = dart_user_add_expr_type(
                    object,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                ) {
                    if operator_return_types.contains_key(&(type_name, "__getitem__".to_string())) {
                        expr.kind = ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new((**object).clone()),
                                field: "__getitem__".to_string(),
                                null_safe: false,
                            })),
                            args: vec![Argument::positional((**index).clone())],
                            optional: false,
                        };
                    }
                }
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_user_add_calls_in_expr(
                value,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            if let Some(call) = dart_user_index_set_call(
                target,
                (**value).clone(),
                env,
                current_class,
                add_return_types,
                operator_return_types,
            ) {
                expr.kind = call.kind;
            } else if let Some(call) = dart_index_set_call(target, (**value).clone()) {
                *expr = call;
                return;
            } else {
                rewrite_user_add_calls_in_expr(
                    target,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                if let Some(key) = &mut elem.key {
                    rewrite_user_add_calls_in_expr(
                        key,
                        env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
                rewrite_user_add_calls_in_expr(
                    &mut elem.value,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_user_add_calls_in_expr(
                    item,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields {
                rewrite_user_add_calls_in_expr(
                    value,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                if let ObjectProperty::KeyValue { value, .. } = prop {
                    rewrite_user_add_calls_in_expr(
                        value,
                        env,
                        current_class,
                        add_return_types,
                        operator_return_types,
                    );
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    // A plain `${x}` where `x` is statically a `double` must
                    // render Dart-style (`2.0`, not `2`). The static type comes
                    // from the same operator/field type tracking used for
                    // operator overloading — no runtime type check.
                    InterpolPart::Expr(value) => {
                        let is_double = matches!(
                            dart_user_add_expr_type(
                                value,
                                env,
                                current_class,
                                add_return_types,
                                operator_return_types,
                            )
                            .as_deref(),
                            Some("double")
                        );
                        rewrite_user_add_calls_in_expr(
                            value,
                            env,
                            current_class,
                            add_return_types,
                            operator_return_types,
                        );
                        if is_double {
                            let inner = value.clone();
                            *value = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__dart_double_to_string")),
                                args: vec![Argument::positional(inner)],
                                optional: false,
                            });
                        }
                    }
                    InterpolPart::Formatted(value, _) => {
                        rewrite_user_add_calls_in_expr(
                            value,
                            env,
                            current_class,
                            add_return_types,
                            operator_return_types,
                        );
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(value) => {
                rewrite_user_add_calls_in_expr(
                    value,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
            LambdaBody::Block(stmts) => {
                let mut lambda_env = env.clone();
                rewrite_user_add_calls_in_stmts(
                    stmts,
                    &mut lambda_env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
        },
        ExprKind::IsType { expr: inner, .. } | ExprKind::Cast { expr: inner, .. } => {
            rewrite_user_add_calls_in_expr(
                inner,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_user_add_calls_in_expr(
                left,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            rewrite_user_add_calls_in_expr(
                right,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
        }
        ExprKind::Match { subject, arms } => {
            rewrite_user_add_calls_in_expr(
                subject,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            );
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for condition in conditions {
                        rewrite_user_add_calls_in_expr(
                            condition,
                            env,
                            current_class,
                            add_return_types,
                            operator_return_types,
                        );
                    }
                }
                rewrite_user_add_calls_in_expr(
                    &mut arm.body,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                );
            }
        }
        _ => {}
    }
}

/// Tags a `dart:io` handle record with which of `File`/`Directory`/`Link` it
/// is. Must match `emitter::io_adapter::DART_IO_KIND_KEY`.
const DART_IO_KIND_KEY: &str = "__dart_io";

/// The `dart:io` filesystem handles. Each is a path plus a kind; the kind is
/// what tells `existsSync` to ask about a file rather than a directory, and
/// what `is Directory` tests against.
fn dart_io_handle_kind(callee: &Expression) -> Option<&'static str> {
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    match name.as_str() {
        "File" => Some("file"),
        "Directory" => Some("directory"),
        "Link" => Some("link"),
        _ => None,
    }
}

/// `{ path: <p>, __dart_io: "file" }` — a plain record, so `.path` is an
/// ordinary field read and the io adapter can pull the path off the receiver
/// without any type inference at the call site.
fn dart_io_handle(kind: &str, path: Expression) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("path"),
            value: path,
        },
        ObjectProperty::KeyValue {
            key: Expression::string(DART_IO_KIND_KEY),
            value: Expression::string(kind),
        },
    ]))
}

/// `(t = receiver, t == null ? null : <use(t)>)` — evaluates the receiver once
/// and performs the access only when it is non-null. Dart's `?.`/`?[]` short-
/// circuit the WHOLE access, so the guard has to wrap the use, not just mark it.
fn dart_null_guarded(
    receiver: Expression,
    use_receiver: impl FnOnce(Expression) -> Expression,
) -> Expression {
    let tmp = nsm_tmp("nullsafe");
    let held = || Expression::ident(&tmp);
    let save = Expression::new(ExprKind::Assign {
        target: Box::new(held()),
        value: Box::new(receiver),
    });
    let guard = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(held()),
        right: Box::new(Expression::null()),
    });
    Expression::new(ExprKind::Sequence(vec![
        save,
        Expression::new(ExprKind::Ternary {
            cond: Box::new(guard),
            then: Box::new(Expression::null()),
            else_: Box::new(use_receiver(held())),
        }),
    ]))
}

/// `FileMode.append` — matched on the enum member however it reached the AST:
/// the walker may have folded it to the string `"FileMode.append"`, or left it
/// as a member access.
fn dart_expr_mentions_file_mode(expr: &Expression, member: &str) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(text)) => text == &format!("FileMode.{member}"),
        ExprKind::Member { object, field, .. } => {
            field == member && matches!(&object.kind, ExprKind::Ident(n) if n == "FileMode")
        }
        ExprKind::Ident(name) => name == &format!("FileMode.{member}"),
        // The spelling the walker actually produces for `FileMode.append`.
        ExprKind::StaticAccess { class, member: m } => {
            matches!(&class.kind, ExprKind::Ident(n) if n == "FileMode")
                && matches!(&m.kind, ExprKind::Ident(n) if n == member)
        }
        _ => false,
    }
}

fn dart_user_binary_operator_method(op: &BinOp) -> Option<&'static str> {
    match op {
        BinOp::Add => Some("operator+"),
        BinOp::Sub => Some("operator-"),
        BinOp::Mul => Some("operator*"),
        BinOp::Div => Some("operator/"),
        BinOp::IDiv => Some("operator~/"),
        BinOp::Mod => Some("operator%"),
        BinOp::Eq | BinOp::StrictEq => Some("__eq__"),
        BinOp::Lt => Some("operator<"),
        BinOp::Gt => Some("operator>"),
        BinOp::LtEq => Some("operator<="),
        BinOp::GtEq => Some("operator>="),
        BinOp::BitAnd => Some("operator&"),
        BinOp::BitOr => Some("operator|"),
        BinOp::BitXor => Some("operator^"),
        BinOp::Shl => Some("operator<<"),
        BinOp::Shr => Some("operator>>"),
        BinOp::UShr => Some("operator>>>"),
        _ => None,
    }
}

fn dart_user_unary_operator_method(op: &UnaryOp) -> Option<&'static str> {
    match op {
        UnaryOp::Neg => Some("operator-@unary"),
        UnaryOp::BitNot => Some("operator~"),
        _ => None,
    }
}

fn dart_user_known_class(
    name: &str,
    add_return_types: &HashMap<String, Option<String>>,
    operator_return_types: &HashMap<(String, String), Option<String>>,
) -> bool {
    add_return_types.contains_key(name)
        || operator_return_types
            .keys()
            .any(|(class_name, _)| class_name == name)
}

fn dart_index_set_call(target: &Expression, value: Expression) -> Option<Expression> {
    let ExprKind::Index { object, index, .. } = &target.kind else {
        return None;
    };
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__dart_index_set")),
        args: vec![
            Argument::positional((**object).clone()),
            Argument::positional((**index).clone()),
            Argument::positional(value),
        ],
        optional: false,
    }))
}

fn dart_user_index_set_call(
    target: &Expression,
    value: Expression,
    env: &HashMap<String, String>,
    current_class: Option<&str>,
    add_return_types: &HashMap<String, Option<String>>,
    operator_return_types: &HashMap<(String, String), Option<String>>,
) -> Option<Expression> {
    let ExprKind::Index { object, index, .. } = &target.kind else {
        return None;
    };
    let type_name = dart_user_add_expr_type(
        object,
        env,
        current_class,
        add_return_types,
        operator_return_types,
    )?;
    if !operator_return_types.contains_key(&(type_name, "__setitem__".to_string())) {
        return None;
    }
    let mut object = (**object).clone();
    let mut index = (**index).clone();
    rewrite_user_add_calls_in_expr(
        &mut object,
        env,
        current_class,
        add_return_types,
        operator_return_types,
    );
    rewrite_user_add_calls_in_expr(
        &mut index,
        env,
        current_class,
        add_return_types,
        operator_return_types,
    );
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: "__setitem__".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(index), Argument::positional(value)],
        optional: false,
    }))
}

fn dart_user_add_expr_type(
    expr: &Expression,
    env: &HashMap<String, String>,
    current_class: Option<&str>,
    add_return_types: &HashMap<String, Option<String>>,
    operator_return_types: &HashMap<(String, String), Option<String>>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::This => current_class.map(str::to_string),
        ExprKind::Ident(name) => env.get(name).cloned(),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name)
                if dart_user_known_class(name, add_return_types, operator_return_types) =>
            {
                Some(name.clone())
            }
            _ => None,
        },
        ExprKind::Member { object, field, .. } => {
            let owner = dart_user_add_expr_type(
                object,
                env,
                current_class,
                add_return_types,
                operator_return_types,
            )?;
            operator_return_types
                .get(&(owner, field.clone()))
                .and_then(|ret| ret.clone())
        }
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name)
                if dart_user_known_class(name, add_return_types, operator_return_types) =>
            {
                Some(name.clone())
            }
            ExprKind::Member { object, field, .. } => {
                let owner = dart_user_add_expr_type(
                    object,
                    env,
                    current_class,
                    add_return_types,
                    operator_return_types,
                )?;
                operator_return_types
                    .get(&(owner, field.clone()))
                    .and_then(|ret| ret.clone())
            }
            _ => None,
        },
        _ => None,
    }
}

fn dart_iter_element_type(
    class_name: &str,
    operator_return_types: &HashMap<(String, String), Option<String>>,
) -> Option<String> {
    operator_return_types
        .get(&(class_name.to_string(), "__dart_iter_element".to_string()))
        .and_then(|ty| ty.clone())
}

fn dart_returned_class_from_body(body: &[Statement]) -> Option<String> {
    for stmt in body {
        match &stmt.kind {
            StmtKind::Return(Some(expr)) | StmtKind::Expr(expr) => {
                if let Some(name) = dart_constructor_expr_type(expr) {
                    return Some(name);
                }
            }
            _ => {}
        }
    }
    None
}

fn dart_constructor_expr_type(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            ExprKind::Member { field, .. } => Some(field.clone()),
            _ => None,
        },
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            ExprKind::Member { field, .. } => Some(field.clone()),
            _ => None,
        },
        _ => None,
    }
}

fn rewrite_lambda_with_param_type(
    expr: &mut Expression,
    param_type: &str,
    env: &HashMap<String, String>,
    current_class: Option<&str>,
    add_return_types: &HashMap<String, Option<String>>,
    operator_return_types: &HashMap<(String, String), Option<String>>,
) {
    let ExprKind::Lambda { params, body, .. } = &mut expr.kind else {
        return;
    };
    let Some(param) = params.first() else {
        return;
    };
    let mut lambda_env = env.clone();
    lambda_env.insert(param.name.clone(), param_type.to_string());
    match body {
        LambdaBody::Expr(value) => rewrite_user_add_calls_in_expr(
            value,
            &lambda_env,
            current_class,
            add_return_types,
            operator_return_types,
        ),
        LambdaBody::Block(stmts) => rewrite_user_add_calls_in_stmts(
            stmts,
            &mut lambda_env,
            current_class,
            add_return_types,
            operator_return_types,
        ),
    }
}

fn rewrite_inherited_instance_member_idents(
    body: &mut Vec<Statement>,
    mixin_names: &std::collections::HashSet<String>,
) {
    let mut class_members_by_name: HashMap<String, Vec<ClassMember>> = HashMap::new();
    let mut class_parents_by_name: HashMap<String, Vec<String>> = HashMap::new();
    for stmt in body.iter() {
        if let StmtKind::ClassDecl {
            name,
            parents,
            members,
            ..
        } = &stmt.kind
        {
            class_members_by_name.insert(name.clone(), members.clone());
            class_parents_by_name.insert(name.clone(), parents.clone());
        }
    }

    for stmt in body.iter_mut() {
        if let StmtKind::ClassDecl {
            name,
            parents,
            members,
            ..
        } = &mut stmt.kind
        {
            if mixin_names.contains(name) || parents.is_empty() {
                continue;
            }
            let extra_members = collect_instance_member_names_for_types(
                parents,
                &class_members_by_name,
                &class_parents_by_name,
            );
            let extra_refs: Vec<&str> = extra_members.iter().map(String::as_str).collect();
            rewrite_instance_member_idents(members, &extra_refs);
        }
    }
}

fn apply_inherited_concrete_members(
    body: &mut Vec<Statement>,
    mixin_names: &std::collections::HashSet<String>,
) {
    let mut classes: HashMap<String, (Vec<String>, Vec<ClassMember>)> = HashMap::new();
    for stmt in body.iter() {
        if let StmtKind::ClassDecl {
            name,
            parents,
            members,
            ..
        } = &stmt.kind
        {
            classes.insert(name.clone(), (parents.clone(), members.clone()));
        }
    }

    for stmt in body.iter_mut() {
        if let StmtKind::ClassDecl {
            name,
            parents,
            members,
            ..
        } = &mut stmt.kind
        {
            if mixin_names.contains(name) || parents.is_empty() {
                continue;
            }
            let mut existing: HashSet<String> = members.iter().filter_map(member_name).collect();
            // A member a MIXIN supplies must not be pre-empted by copying the
            // superclass's version in here: Dart says the mixin wins over the
            // base (`class C extends Base with M` → M.greet, not Base.greet).
            // The shared augmentation pass folds mixin members later, and it
            // correctly refuses to overwrite a member the class already has —
            // so an inherited copy landing first would silently win.
            for mixin in dart_class_mixins(name) {
                if let Some((_, mixin_members)) = classes.get(&mixin) {
                    existing.extend(mixin_members.iter().filter_map(member_name));
                }
            }
            let mut inherited = Vec::new();
            for parent in parents.iter() {
                collect_inherited_concrete_members(
                    parent,
                    &classes,
                    &mut HashSet::new(),
                    &mut inherited,
                );
            }
            for member in inherited {
                if let Some(name) = member_name(&member) {
                    if existing.insert(name) {
                        members.push(member);
                    }
                }
            }
        }
    }
}

fn collect_inherited_concrete_members(
    class_name: &str,
    classes: &HashMap<String, (Vec<String>, Vec<ClassMember>)>,
    seen: &mut HashSet<String>,
    out: &mut Vec<ClassMember>,
) {
    if !seen.insert(class_name.to_string()) {
        return;
    }
    let Some((parents, members)) = classes.get(class_name) else {
        return;
    };
    for parent in parents {
        collect_inherited_concrete_members(parent, classes, seen, out);
    }
    for member in members {
        if let Some(member) = inheritable_concrete_member(member) {
            out.push(member);
        }
    }
}

fn inheritable_concrete_member(member: &ClassMember) -> Option<ClassMember> {
    match member {
        ClassMember::Method(stmt) => match &stmt.kind {
            StmtKind::FunctionDecl {
                body, modifiers, ..
            } if !modifiers.is_static && !modifiers.is_abstract && !body.is_empty() => {
                Some(member.clone())
            }
            _ => None,
        },
        ClassMember::Property {
            getter,
            setter,
            modifiers,
            ..
        } if !modifiers.is_static && (getter.is_some() || setter.is_some()) => Some(member.clone()),
        _ => None,
    }
}

/// Dart falls through an EMPTY case body and only an empty one — verified
/// against the SDK:
///
/// ```dart
/// switch (10) { case 10: case 20: print('tens'); break; }   // prints tens
/// switch (5)  { case 5: default: print('via-default'); }    // prints via-default
/// ```
///
/// A case that DOES have a body breaks implicitly; Dart 3 rejects falling out
/// of one. So this is not JS's `switch_fallthrough` (which would also chain
/// non-empty bodies) — it is a syntactic grouping, and it normalizes away here:
///
/// - an empty case's conditions join the NEXT case's condition list;
/// - an empty case with no following case is DROPPED, so its value matches
///   nothing and reaches `default` — which is exactly where Dart sends it.
///
/// Without this an empty case matched, did nothing, and exited: `case 10:`
/// above swallowed the value and printed nothing at all.
fn merge_empty_fallthrough_cases(cases: &mut Vec<SwitchCase>) {
    let mut pending: Vec<CaseCondition> = Vec::new();
    let mut merged: Vec<SwitchCase> = Vec::new();
    for case in cases.drain(..) {
        // `conditions: vec![]` IS the default arm — never merge into it, or a
        // value would stop reaching the arms after it.
        if case.conditions.is_empty() {
            merged.push(case);
            continue;
        }
        if case.body.is_empty() {
            pending.extend(case.conditions);
            continue;
        }
        let mut case = case;
        if !pending.is_empty() {
            let mut conditions = std::mem::take(&mut pending);
            conditions.append(&mut case.conditions);
            case.conditions = conditions;
        }
        merged.push(case);
    }
    // Trailing empty cases are dropped on purpose: unmatched reaches `default`.
    *cases = merged;
}

/// Backing storage for a field that overrides an inherited getter.
fn dart_override_storage_name(field: &str) -> String {
    format!("__dart_ovr_{field}")
}

/// In Dart a subclass FIELD overrides an inherited GETTER — `class Toggle {
/// Object get state => false; }` / `class Sub extends Toggle { bool state =
/// true; }` reads `true`. Emitted naively that is a plain `this.state = true`
/// against an accessor the parent installed with no setter, so the write is
/// dropped and the read still runs the parent's getter: every such program
/// silently returned the SUPERCLASS's value.
///
/// So the field is re-expressed as what it means — a property override backed
/// by its own storage. That is ordinary walker normalisation: the class-level
/// machinery then installs the accessor on the subclass, which shadows the
/// parent's the way any override does, and no shared code changes.
///
/// Only fires when an ancestor really declares that name as a property; a
/// field with no inherited counterpart stays a plain field.
fn override_inherited_getter_fields(body: &mut [Statement]) {
    let mut properties: HashMap<String, HashSet<String>> = HashMap::new();
    let mut parents_of: HashMap<String, Vec<String>> = HashMap::new();
    for stmt in body.iter() {
        if let StmtKind::ClassDecl {
            name,
            parents,
            members,
            ..
        } = &stmt.kind
        {
            parents_of.insert(name.clone(), parents.clone());
            properties.insert(
                name.clone(),
                members
                    .iter()
                    .filter_map(|m| match m {
                        ClassMember::Property {
                            name,
                            getter: Some(_),
                            modifiers,
                            ..
                        } if !modifiers.is_static => Some(name.clone()),
                        _ => None,
                    })
                    .collect(),
            );
        }
    }

    for stmt in body.iter_mut() {
        let StmtKind::ClassDecl {
            name,
            parents,
            members,
            ..
        } = &mut stmt.kind
        else {
            continue;
        };
        // Ancestor property names, walking the chain. `seen` keeps a cyclic
        // `extends` from looping.
        let mut inherited: HashSet<String> = HashSet::new();
        let mut queue: Vec<String> = parents.clone();
        let mut seen: HashSet<String> = HashSet::from([name.clone()]);
        while let Some(ancestor) = queue.pop() {
            if !seen.insert(ancestor.clone()) {
                continue;
            }
            if let Some(names) = properties.get(&ancestor) {
                inherited.extend(names.iter().cloned());
            }
            if let Some(grandparents) = parents_of.get(&ancestor) {
                queue.extend(grandparents.iter().cloned());
            }
        }
        if inherited.is_empty() {
            continue;
        }

        let mut rewritten = Vec::new();
        for member in members.iter() {
            let ClassMember::Field {
                name: fname,
                type_hint,
                init,
                modifiers,
                ..
            } = member
            else {
                continue;
            };
            if modifiers.is_static || !inherited.contains(fname) {
                continue;
            }
            rewritten.push((
                fname.clone(),
                type_hint.clone(),
                init.clone(),
                modifiers.clone(),
            ));
        }
        if rewritten.is_empty() {
            continue;
        }
        let replaced: HashSet<String> = rewritten.iter().map(|(n, ..)| n.clone()).collect();
        members
            .retain(|m| !matches!(m, ClassMember::Field { name, .. } if replaced.contains(name)));

        for (fname, type_hint, init, modifiers) in rewritten {
            let storage = dart_override_storage_name(&fname);
            let storage_ref = || {
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: storage.clone(),
                    null_safe: false,
                })
            };
            members.push(ClassMember::Field {
                name: storage.clone(),
                type_hint: type_hint.clone(),
                init,
                modifiers: modifiers.clone(),
                with_events: false,
                array_bounds: None,
            });
            let value_param = Param {
                name: "__dart_ovr_value".to_string(),
                type_hint: type_hint.clone().map(Into::into),
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            };
            members.push(ClassMember::Property {
                name: fname,
                type_hint,
                getter: Some(vec![Statement::new(StmtKind::Return(Some(storage_ref())))]),
                setter: Some(PropertySetter {
                    param: value_param.clone(),
                    body: vec![Statement::new(StmtKind::Expr(Expression::new(
                        ExprKind::Assign {
                            target: Box::new(storage_ref()),
                            value: Box::new(Expression::ident(&value_param.name)),
                        },
                    )))],
                }),
                is_auto: false,
                modifiers,
            });
        }
    }
}

fn member_name(m: &ClassMember) -> Option<String> {
    match m {
        ClassMember::Field { name, .. } => Some(name.clone()),
        ClassMember::Method(stmt) => match &stmt.kind {
            StmtKind::FunctionDecl { name, .. } => Some(name.clone()),
            _ => None,
        },
        ClassMember::Const { name, .. } => Some(name.clone()),
        ClassMember::Property { name, .. } => Some(name.clone()),
        _ => None,
    }
}

fn instance_member_name(m: &ClassMember) -> Option<String> {
    match m {
        ClassMember::Field {
            name, modifiers, ..
        } if !modifiers.is_static => Some(name.clone()),
        ClassMember::Method(stmt) => match &stmt.kind {
            StmtKind::FunctionDecl {
                name, modifiers, ..
            } if !modifiers.is_static => Some(name.clone()),
            _ => None,
        },
        ClassMember::Property {
            name, modifiers, ..
        } if !modifiers.is_static => Some(name.clone()),
        _ => None,
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Top-level items
// ════════════════════════════════════════════════════════════════════════════

fn walk_top_level(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::class_declaration => walk_class_decl(pair)?,
        Rule::mixin_declaration => walk_mixin_decl(pair)?,
        Rule::extension_type_declaration => walk_extension_type_decl(pair)?,
        Rule::extension_declaration => walk_extension_decl(pair)?,
        Rule::enum_declaration => walk_enum_decl(pair)?,
        Rule::typedef_declaration => return Ok(None), // type aliases are discarded
        Rule::function_declaration => walk_function_decl(pair)?,
        Rule::variable_declaration_statement => walk_var_decl_stmt(pair)?,
        Rule::expression_statement => {
            let expr = walk_expression(pair.into_inner().next().ok_or("empty expr stmt")?)?;
            StmtKind::Expr(expr)
        }
        Rule::annotation => return Ok(None), // annotations discarded at top level
        _ => return Ok(None),
    };
    Ok(Some(Statement::with_span(kind, span)))
}

// ════════════════════════════════════════════════════════════════════════════
// Imports
// ════════════════════════════════════════════════════════════════════════════

fn walk_import(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut path = String::new();
    let mut alias: Option<String> = None;
    let mut show_names: Vec<ImportName> = Vec::new();
    let mut hide_names: Vec<String> = Vec::new();
    let mut _deferred = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::string_literal => path = unquote_string_literal(&p),
            Rule::as_clause => {
                for c in p.into_inner() {
                    if c.as_rule() == Rule::ident_name {
                        alias = Some(c.as_str().to_string());
                    }
                }
            }
            Rule::deferred_clause => {
                _deferred = true;
                for c in p.into_inner() {
                    if c.as_rule() == Rule::as_clause {
                        for a in c.into_inner() {
                            if a.as_rule() == Rule::ident_name {
                                alias = Some(a.as_str().to_string());
                            }
                        }
                    }
                }
            }
            Rule::show_clause => {
                for c in p.into_inner() {
                    if c.as_rule() == Rule::ident_list {
                        for name_pair in c.into_inner() {
                            if name_pair.as_rule() == Rule::ident_name {
                                show_names.push(ImportName {
                                    name: name_pair.as_str().to_string(),
                                    alias: None,
                                });
                            }
                        }
                    }
                }
            }
            Rule::hide_clause => {
                for c in p.into_inner() {
                    if c.as_rule() == Rule::ident_list {
                        for name_pair in c.into_inner() {
                            if name_pair.as_rule() == Rule::ident_name {
                                hide_names.push(name_pair.as_str().to_string());
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }

    let kind = if !show_names.is_empty() {
        ImportKind::Named {
            path,
            names: show_names,
            level: 0,
        }
    } else {
        ImportKind::Simple { path, alias }
    };

    Ok(Import { kind, span })
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::empty_statement => StmtKind::Empty,

        Rule::block_statement => {
            let mut stmts = Vec::new();
            for p in pair.into_inner() {
                if let Some(s) = walk_statement(p)? {
                    stmts.push(s);
                }
            }
            StmtKind::Block(stmts)
        }

        Rule::variable_declaration_statement => walk_var_decl_stmt(pair)?,

        Rule::if_statement => walk_if(pair)?,

        Rule::for_statement => walk_for(pair)?,

        Rule::while_statement => walk_while(pair)?,

        Rule::do_while_statement => walk_do_while(pair)?,

        Rule::switch_statement => walk_switch(pair)?,

        Rule::return_statement => {
            let expr = pair
                .into_inner()
                .find(|p| !is_kw(p.as_rule()))
                .map(walk_expression)
                .transpose()?;
            StmtKind::Return(expr)
        }

        Rule::break_statement => {
            let label = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string());
            StmtKind::Break(match label {
                Some(l) => BreakTarget::Label(l),
                None => BreakTarget::Implicit,
            })
        }

        Rule::continue_statement => {
            let label = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string());
            StmtKind::Continue(match label {
                Some(l) => ContinueTarget::Label(l),
                None => ContinueTarget::Implicit,
            })
        }

        Rule::labeled_statement => {
            let mut inner = pair.into_inner();
            let label = inner
                .next()
                .ok_or("labeled statement: missing label")?
                .as_str()
                .to_string();
            let body = walk_statement(inner.next().ok_or("labeled statement: missing body")?)?
                .ok_or("labeled statement: empty body")?;
            StmtKind::Labeled {
                label,
                body: Box::new(body),
            }
        }

        Rule::throw_statement => {
            let inner = pair.into_inner().next().ok_or("throw: missing expr")?;
            let expr = walk_expression(inner)?;
            StmtKind::Throw {
                expr: Some(expr),
                cause: None,
            }
        }

        Rule::yield_statement => walk_yield_statement(pair)?,

        Rule::rethrow_statement => StmtKind::Throw {
            expr: None,
            cause: None,
        },

        Rule::try_statement => walk_try(pair)?,

        Rule::assert_statement => {
            let mut exprs: Vec<Expression> = Vec::new();
            for p in pair.into_inner() {
                if !is_kw(p.as_rule()) {
                    exprs.push(walk_expression(p)?);
                }
            }
            let test = exprs.remove(0);
            let msg = if exprs.is_empty() {
                None
            } else {
                Some(exprs.remove(0))
            };
            StmtKind::Assert { test, msg }
        }

        Rule::function_declaration => walk_function_decl(pair)?,

        Rule::expression_statement => {
            let inner = pair.into_inner().next().ok_or("empty expr stmt")?;
            let expr = walk_expression(inner)?;
            StmtKind::Expr(expr)
        }

        Rule::class_declaration => walk_class_decl(pair)?,
        Rule::enum_declaration => walk_enum_decl(pair)?,

        _ => return Ok(None),
    };
    Ok(Some(Statement::with_span(kind, span)))
}

fn walk_statement_into_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    if matches!(
        pair.as_rule(),
        Rule::block_statement | Rule::function_body_block
    ) {
        let mut stmts = Vec::new();
        for p in pair.into_inner() {
            if let Some(s) = walk_statement(p)? {
                stmts.push(s);
            }
        }
        Ok(stmts)
    } else {
        match walk_statement(pair)? {
            Some(s) => Ok(vec![s]),
            None => Ok(Vec::new()),
        }
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Variable declarations
// ════════════════════════════════════════════════════════════════════════════

fn walk_var_decl_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // variable_declaration_statement = {
    //     var_modifiers ~ type_or_var ~ var_declarator ~ ("," ~ var_declarator)* ~ ";"
    // }
    let mut var_kind = VarDeclKind::Let;
    let mut declarations = Vec::new();
    let mut type_hint: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_modifiers => {
                let txt = p.as_str().trim();
                if txt.contains("final") || txt.contains("const") {
                    var_kind = VarDeclKind::Const;
                }
            }
            Rule::type_or_var => {
                let inner_text = p.as_str().trim();
                if inner_text != "var" {
                    // It's a type annotation, not bare `var`
                    // Check inner children for var_kw
                    let has_var_kw = p.clone().into_inner().any(|c| c.as_rule() == Rule::var_kw);
                    if !has_var_kw {
                        type_hint = Some(inner_text.to_string());
                    }
                }
            }
            Rule::typed_var_declarator => {
                let decl = walk_var_declarator(p, type_hint.clone())?;
                declarations.push(decl);
            }
            Rule::var_declarator => {
                if let Some(block) = lower_destructuring_var_declarator(
                    p.clone(),
                    var_kind.clone(),
                    declarations.len(),
                )? {
                    if !declarations.is_empty() {
                        let mut body = vec![Statement::new(StmtKind::VarDecl {
                            declarations,
                            kind: var_kind,
                        })];
                        body.extend(block);
                        return Ok(StmtKind::Block(body));
                    }
                    return Ok(StmtKind::Block(block));
                } else {
                    let decl = walk_var_declarator(p, type_hint.clone())?;
                    declarations.push(decl);
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::VarDecl {
        declarations,
        kind: var_kind,
    })
}

fn walk_var_declarator(
    pair: Pair<Rule>,
    type_hint: Option<String>,
) -> Result<VarDeclarator, String> {
    let mut name = String::new();
    let mut init = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::assignment_expression => init = Some(walk_expression(p)?),
            _ => {
                if init.is_none() {
                    init = Some(walk_expression(p)?);
                }
            }
        }
    }

    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint: type_hint.map(Into::into),
        init,
        array_bounds: None,
        with_events: false,
    })
}

fn lower_destructuring_var_declarator(
    pair: Pair<Rule>,
    kind: VarDeclKind,
    ordinal: usize,
) -> Result<Option<Vec<Statement>>, String> {
    let span = to_span(&pair);
    let mut pattern = None;
    let mut init = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::destructuring_pattern => {
                pattern = p.into_inner().next();
            }
            Rule::assignment_expression => init = Some(walk_expression(p)?),
            _ => {}
        }
    }
    let Some(pattern) = pattern else {
        return Ok(None);
    };
    let tmp = format!(
        "__dart_destructure_{}_{}_{}",
        span.start_line, span.start_col, ordinal
    );
    let subject = Expression::ident(&tmp);
    let bindings = dart_decl_pattern_bindings(pattern, &subject)?;
    let mut body = Vec::new();
    body.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(tmp),
            type_hint: None,
            init: Some(init.unwrap_or_else(Expression::null)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    }));
    for (name, value) in bindings {
        body.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(name),
                type_hint: None,
                init: Some(value),
                array_bounds: None,
                with_events: false,
            }],
            kind: kind.clone(),
        }));
    }
    Ok(Some(body))
}

fn lower_for_in_header_parts(
    header_inner: Pair<Rule>,
    mut body: Vec<Statement>,
) -> Result<(String, Expression, Vec<Statement>), String> {
    let span = to_span(&header_inner);
    let mut var_name = String::new();
    let mut iter_expr = None;
    let mut pattern = None;

    for p in header_inner.into_inner() {
        match p.as_rule() {
            Rule::final_kw | Rule::var_kw | Rule::type_annotation => {}
            Rule::destructuring_pattern => pattern = p.into_inner().next(),
            Rule::ident_name => var_name = p.as_str().to_string(),
            _ => iter_expr = Some(walk_expression(p)?),
        }
    }

    if let Some(pattern) = pattern {
        var_name = format!("__dart_for_in_{}_{}", span.start_line, span.start_col);
        let subject = Expression::ident(&var_name);
        let bindings = dart_decl_pattern_bindings(pattern, &subject)?;
        let mut prefix = Vec::new();
        for (name, value) in bindings {
            prefix.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: None,
                    init: Some(value),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }));
        }
        prefix.extend(body);
        body = prefix;
    }

    let iter = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__dart_for_in_iterable")),
        args: vec![Argument::positional(
            iter_expr.ok_or("for-in: missing iterable")?,
        )],
        optional: false,
    });

    Ok((var_name, iter, body))
}

fn dart_decl_pattern_bindings(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<Vec<(String, Expression)>, String> {
    match pair.as_rule() {
        Rule::destructuring_pattern | Rule::pattern | Rule::primary_pattern => {
            let mut out = Vec::new();
            for child in pair.into_inner() {
                out.extend(dart_decl_pattern_bindings(child, subject)?);
            }
            Ok(out)
        }
        Rule::record_pattern => {
            let is_grouping_pattern = !pair.as_str().contains(',');
            let mut out = Vec::new();
            let mut index = 0usize;
            let fields: Vec<Pair<Rule>> = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::record_pattern_field)
                .collect();
            if is_grouping_pattern && fields.len() == 1 {
                let children: Vec<Pair<Rule>> = fields[0].clone().into_inner().collect();
                if children.len() == 1 {
                    return dart_decl_pattern_bindings(children[0].clone(), subject);
                }
            }
            for field in fields {
                let children: Vec<Pair<Rule>> = field.into_inner().collect();
                let (target, pat) = if children.len() == 2 {
                    (
                        Expression::new(ExprKind::Member {
                            object: Box::new(subject.clone()),
                            field: children[0].as_str().to_string(),
                            null_safe: false,
                        }),
                        children[1].clone(),
                    )
                } else {
                    let target = Expression::new(ExprKind::Index {
                        object: Box::new(subject.clone()),
                        index: Box::new(Expression::int(index as i64)),
                        null_safe: false,
                    });
                    index += 1;
                    (target, children[0].clone())
                };
                out.extend(dart_decl_pattern_bindings(pat, &target)?);
            }
            Ok(out)
        }
        Rule::list_pattern => {
            let mut out = Vec::new();
            let mut index = 0usize;
            for elem in pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::list_pattern_element)
            {
                let child = elem.into_inner().next().ok_or("list pattern: empty")?;
                if child.as_rule() == Rule::rest_pattern {
                    if let Some(name) = child
                        .into_inner()
                        .find(|p| p.as_rule() == Rule::ident_name)
                        .map(|p| p.as_str().to_string())
                    {
                        if name != "_" {
                            out.push((
                                name,
                                dart_method_call(
                                    subject.clone(),
                                    "sublist",
                                    vec![Expression::int(index as i64)],
                                ),
                            ));
                        }
                    }
                    continue;
                }
                let target = Expression::new(ExprKind::Index {
                    object: Box::new(subject.clone()),
                    index: Box::new(Expression::int(index as i64)),
                    null_safe: false,
                });
                out.extend(dart_decl_pattern_bindings(child, &target)?);
                index += 1;
            }
            Ok(out)
        }
        Rule::map_pattern => {
            let mut out = Vec::new();
            for entry in pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::map_pattern_entry)
            {
                let mut inner = entry.into_inner();
                let key = walk_expression(inner.next().ok_or("map pattern: missing key")?)?;
                let pat = inner.next().ok_or("map pattern: missing value")?;
                let target = Expression::new(ExprKind::Index {
                    object: Box::new(subject.clone()),
                    index: Box::new(key),
                    null_safe: false,
                });
                out.extend(dart_decl_pattern_bindings(pat, &target)?);
            }
            Ok(out)
        }
        Rule::object_pattern => {
            let mut out = Vec::new();
            for field in pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::object_pattern_field)
            {
                let mut inner = field.into_inner();
                let name = inner
                    .next()
                    .ok_or("object pattern: missing field")?
                    .as_str()
                    .to_string();
                let pat = inner.next().ok_or("object pattern: missing pattern")?;
                let target = Expression::new(ExprKind::Member {
                    object: Box::new(subject.clone()),
                    field: name,
                    null_safe: false,
                });
                out.extend(dart_decl_pattern_bindings(pat, &target)?);
            }
            Ok(out)
        }
        Rule::variable_pattern => {
            let name = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string());
            Ok(name
                .filter(|name| name != "_")
                .map(|name| vec![(name, subject.clone())])
                .unwrap_or_default())
        }
        Rule::constant_pattern => {
            let children: Vec<Pair<Rule>> = pair.into_inner().collect();
            if children.len() == 1 && children[0].as_rule() == Rule::ident_name {
                let name = children[0].as_str().to_string();
                if name != "_" {
                    return Ok(vec![(name, subject.clone())]);
                }
            }
            Ok(Vec::new())
        }
        Rule::wildcard_pattern => Ok(Vec::new()),
        _ => Ok(Vec::new()),
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Function declarations
// ════════════════════════════════════════════════════════════════════════════

fn consume_dart_type_params(pair: Pair<Rule>) {
    let _ = common_generics::parse_generic_params_hint(pair.as_str());
}

fn walk_function_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut return_type: Option<String> = None;
    let mut is_async = false;
    let mut is_generator = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_annotation => {
                if name.is_empty() {
                    // Return type comes before name
                    return_type = Some(p.as_str().to_string());
                }
            }
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::type_params => consume_dart_type_params(p),
            Rule::param_list => params = walk_params(p)?,
            Rule::async_kw => is_async = true,
            Rule::generator_marker => is_generator = true,
            Rule::function_body => body = walk_function_body(p)?,
            _ => {}
        }
    }

    if is_async && !is_generator && body.is_empty() {
        body.push(Statement::new(StmtKind::Return(Some(dart_future_value(
            Expression::null(),
        )))));
    }
    is_generator = is_generator || body_has_yield(&body);

    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async,
        is_generator,
        is_sub: false,
    })
}

fn walk_function_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    // function_body = { arrow_body | function_body_block | empty_body }
    let inner = pair.into_inner().next();
    match inner {
        None => Ok(Vec::new()),
        Some(p) => match p.as_rule() {
            Rule::arrow_body => {
                // arrow_body = { "=>" ~ expression ~ ";" }
                let expr_pair = p.into_inner().next().ok_or("arrow body: no expr")?;
                if expr_pair.as_rule() == Rule::throw_expression {
                    let thrown = walk_throw_expression(expr_pair)?;
                    return Ok(vec![Statement::new(StmtKind::Throw {
                        expr: Some(thrown),
                        cause: None,
                    })]);
                }
                let expr = walk_expression(expr_pair)?;
                Ok(vec![Statement::new(StmtKind::Return(Some(expr)))])
            }
            Rule::function_body_block => walk_statement_into_body(p),
            Rule::empty_body => Ok(Vec::new()),
            _ => Ok(Vec::new()),
        },
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Parameters
// ════════════════════════════════════════════════════════════════════════════

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_group => {
                for inner in p.into_inner() {
                    match inner.as_rule() {
                        Rule::param => out.push(walk_param(inner)?),
                        Rule::optional_positional_params => {
                            for op in inner.into_inner() {
                                if op.as_rule() == Rule::param {
                                    let mut param = walk_param(op)?;
                                    param.is_optional = true;
                                    out.push(param);
                                }
                            }
                        }
                        Rule::named_params => {
                            for np in inner.into_inner() {
                                if np.as_rule() == Rule::param {
                                    let mut param = walk_param(np)?;
                                    param.is_optional = true;
                                    out.push(param);
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Rule::param => out.push(walk_param(p)?),
            _ => {}
        }
    }
    Ok(out)
}

fn walk_lambda_param_pair(lparam: Pair<Rule>) -> Result<Param, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut default = None;
    for inner in lparam.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => name = inner.as_str().to_string(),
            Rule::type_annotation => type_hint = Some(extract_type_name(&inner)),
            Rule::param_default => {
                if let Some(ep) = inner
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::assignment_expression)
                {
                    default = Some(walk_expression(ep)?);
                }
            }
            Rule::typed_lambda_param => {
                for ti in inner.into_inner() {
                    match ti.as_rule() {
                        Rule::ident_name => name = ti.as_str().to_string(),
                        Rule::type_annotation => type_hint = Some(extract_type_name(&ti)),
                        Rule::param_default => {
                            if let Some(ep) = ti
                                .into_inner()
                                .find(|c| c.as_rule() == Rule::assignment_expression)
                            {
                                default = Some(walk_expression(ep)?);
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }
    let is_optional = default.is_some();
    Ok(Param {
        name,
        type_hint: type_hint.map(Into::into),
        default,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional,
        is_nullable: false,
    })
}

fn walk_param(pair: Pair<Rule>) -> Result<Param, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut default: Option<Expression> = None;
    let mut is_this_param = false;
    let mut is_super_param = false;

    fn handle(
        p: Pair<Rule>,
        name: &mut String,
        type_hint: &mut Option<String>,
        default: &mut Option<Expression>,
        is_this: &mut bool,
        is_super: &mut bool,
    ) -> Result<(), String> {
        match p.as_rule() {
            Rule::required_kw | Rule::covariant_kw | Rule::final_kw => {}
            Rule::this_param_prefix => *is_this = true,
            Rule::super_param_prefix => *is_super = true,
            Rule::type_annotation => *type_hint = Some(extract_type_name(&p)),
            Rule::ident_name => *name = p.as_str().to_string(),
            Rule::this_param | Rule::super_param | Rule::typed_or_untyped_param => {
                if p.as_rule() == Rule::this_param {
                    *is_this = true;
                } else if p.as_rule() == Rule::super_param {
                    *is_super = true;
                }
                // Unwrap the wrapper rule and recurse into its children.
                for inner in p.into_inner() {
                    handle(inner, name, type_hint, default, is_this, is_super)?;
                }
            }
            Rule::param_default => {
                let expr_pair = p
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::assignment_expression);
                if let Some(ep) = expr_pair {
                    *default = Some(walk_expression(ep)?);
                }
            }
            _ => {}
        }
        Ok(())
    }
    for p in pair.into_inner() {
        handle(
            p,
            &mut name,
            &mut type_hint,
            &mut default,
            &mut is_this_param,
            &mut is_super_param,
        )?;
    }

    // this.field params: we keep the bare name. The constructor walker
    // will synthesise `this.name = name;` assignments.
    let _ = (is_this_param, is_super_param); // info consumed by constructor_declaration walker

    let is_optional = default.is_some();
    Ok(Param {
        name,
        type_hint: type_hint.map(Into::into),
        default,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional,
        is_nullable: false,
    })
}

/// Walk a param and also return whether it was a `this.x` or `super.x` param.
fn walk_param_with_this(pair: Pair<Rule>) -> Result<(Param, bool, bool), String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut default: Option<Expression> = None;
    let mut is_this_param = false;
    let mut is_super_param = false;

    fn handle(
        p: Pair<Rule>,
        name: &mut String,
        type_hint: &mut Option<String>,
        default: &mut Option<Expression>,
        is_this: &mut bool,
        is_super: &mut bool,
    ) -> Result<(), String> {
        match p.as_rule() {
            Rule::required_kw | Rule::covariant_kw | Rule::final_kw => {}
            Rule::this_param_prefix => *is_this = true,
            Rule::super_param_prefix => *is_super = true,
            Rule::type_annotation => *type_hint = Some(extract_type_name(&p)),
            Rule::ident_name => *name = p.as_str().to_string(),
            Rule::this_param | Rule::super_param | Rule::typed_or_untyped_param => {
                if p.as_rule() == Rule::this_param {
                    *is_this = true;
                } else if p.as_rule() == Rule::super_param {
                    *is_super = true;
                }
                for inner in p.into_inner() {
                    handle(inner, name, type_hint, default, is_this, is_super)?;
                }
            }
            Rule::param_default => {
                let expr_pair = p
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::assignment_expression);
                if let Some(ep) = expr_pair {
                    *default = Some(walk_expression(ep)?);
                }
            }
            _ => {}
        }
        Ok(())
    }
    for p in pair.into_inner() {
        handle(
            p,
            &mut name,
            &mut type_hint,
            &mut default,
            &mut is_this_param,
            &mut is_super_param,
        )?;
    }

    let is_optional = default.is_some();
    let param = Param {
        name,
        type_hint: type_hint.map(Into::into),
        default,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional,
        is_nullable: false,
    };
    Ok((param, is_this_param, is_super_param))
}

// ════════════════════════════════════════════════════════════════════════════
// Class declarations
// ════════════════════════════════════════════════════════════════════════════

fn walk_class_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();
    let mut modifiers = ClassModifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::abstract_kw => modifiers.is_abstract = true,
            Rule::class_modifier => {
                if p.into_inner().any(|m| m.as_rule() == Rule::abstract_kw) {
                    modifiers.is_abstract = true;
                }
            }
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::type_params => consume_dart_type_params(p),
            Rule::extends_clause => {
                if let Some(type_name) = extract_type_name_from_clause(&p) {
                    parents.push(type_name);
                }
            }
            Rule::with_clause => {
                // Mixins become additional parents
                for ta in p.into_inner() {
                    if ta.as_rule() == Rule::type_annotation_list {
                        for t in ta.into_inner() {
                            if t.as_rule() == Rule::type_annotation {
                                parents.push(extract_type_name(&t));
                            }
                        }
                    }
                }
            }
            Rule::implements_clause => {
                for ta in p.into_inner() {
                    if ta.as_rule() == Rule::type_annotation_list {
                        for t in ta.into_inner() {
                            if t.as_rule() == Rule::type_annotation {
                                interfaces.push(extract_type_name(&t));
                            }
                        }
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    match m.as_rule() {
                        Rule::constructor_declaration
                        | Rule::operator_declaration
                        | Rule::getter_declaration
                        | Rule::setter_declaration
                        | Rule::method_declaration
                        | Rule::field_declaration => {
                            if let Some(member) = walk_class_member(m, &name)? {
                                members.push(member);
                            }
                        }
                        Rule::annotation => {} // annotations on members — discard
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    // Static-field rewrite: in static methods, bare `count` (matching
    // a static field name) means `ClassName.count`. Walker rewrites
    // here so the shared compiler doesn't need to track method
    // staticity. Same idea as Fortran's walker normalizing language
    // idioms before they hit the common AST.
    let static_field_names: Vec<String> = members
        .iter()
        .filter_map(|m| {
            if let ClassMember::Field {
                name: fname,
                modifiers,
                ..
            } = m
            {
                if modifiers.is_static {
                    Some(fname.clone())
                } else {
                    None
                }
            } else {
                None
            }
        })
        .collect();
    if !static_field_names.is_empty() {
        for member in members.iter_mut() {
            if let ClassMember::Method(stmt) = member {
                if let StmtKind::FunctionDecl {
                    body, modifiers, ..
                } = &mut stmt.kind
                {
                    if modifiers.is_static {
                        for s in body.iter_mut() {
                            rewrite_static_idents(s, &name, &static_field_names);
                        }
                    }
                }
            }
        }
    }
    rewrite_instance_member_idents(&mut members, &[]);

    let interfaces = with_catalog_ancestry(&parents, interfaces);
    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers,
        decorators: vec![],
    })
}

/// Carry a catalog base class's ancestry onto a user class that extends it.
///
/// `class MyWidget extends StatelessWidget` must satisfy `is Widget` too, but
/// the catalog's abstract chain (`StatelessWidget → Widget`) lives in the
/// namespace tree, not in the compiler's class table — so the user class would
/// only ever know its immediate parent. Adding the base's remaining ancestry as
/// interfaces gives the object the same `__types` membership it would have had
/// if the whole chain were declared, which is what `is` tests.
fn with_catalog_ancestry(parents: &[String], mut interfaces: Vec<String>) -> Vec<String> {
    for parent in parents {
        if USER_DECLARED_TYPES.with(|s| s.borrow().contains(parent)) {
            continue;
        }
        let Some(class) = vybe_platform_flutter::emitter::flutter_classes()
            .iter()
            .find(|c| c.name == *parent)
        else {
            continue;
        };
        // `ancestry` is self-first; the parent itself is already the declared
        // superclass, so only its ancestors are added.
        for ancestor in vybe_platform_flutter::emitter::catalog::ancestry(class)
            .into_iter()
            .skip(1)
        {
            if !interfaces.iter().any(|i| *i == ancestor) {
                interfaces.push(ancestor);
            }
        }
    }
    interfaces
}

/// Rewrite bare `field` to `ClassName.field` for matching static fields.
/// Recursive walk of every Statement / Expression in a static method body.
fn rewrite_static_idents(stmt: &mut Statement, class_name: &str, static_fields: &[String]) {
    match &mut stmt.kind {
        StmtKind::Expr(e) | StmtKind::Return(Some(e)) => {
            rewrite_static_idents_expr(e, class_name, static_fields)
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations.iter_mut() {
                if let Some(init) = &mut d.init {
                    rewrite_static_idents_expr(init, class_name, static_fields);
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
            ..
        } => {
            rewrite_static_idents_expr(cond, class_name, static_fields);
            for s in then_body.iter_mut() {
                rewrite_static_idents(s, class_name, static_fields);
            }
            for (c, body) in elifs.iter_mut() {
                rewrite_static_idents_expr(c, class_name, static_fields);
                for s in body.iter_mut() {
                    rewrite_static_idents(s, class_name, static_fields);
                }
            }
            if let Some(body) = else_body {
                for s in body.iter_mut() {
                    rewrite_static_idents(s, class_name, static_fields);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            if let Some(i) = init.as_deref_mut() {
                rewrite_static_idents(i, class_name, static_fields);
            }
            if let Some(c) = cond {
                rewrite_static_idents_expr(c, class_name, static_fields);
            }
            if let Some(u) = update {
                rewrite_static_idents_expr(u, class_name, static_fields);
            }
            for s in body.iter_mut() {
                rewrite_static_idents(s, class_name, static_fields);
            }
        }
        StmtKind::ForIn { iter, body, .. } => {
            rewrite_static_idents_expr(iter, class_name, static_fields);
            for s in body.iter_mut() {
                rewrite_static_idents(s, class_name, static_fields);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            rewrite_static_idents_expr(cond, class_name, static_fields);
            for s in body.iter_mut() {
                rewrite_static_idents(s, class_name, static_fields);
            }
        }
        StmtKind::Block(stmts) => {
            for s in stmts.iter_mut() {
                rewrite_static_idents(s, class_name, static_fields);
            }
        }
        _ => {}
    }
}

fn rewrite_static_idents_expr(expr: &mut Expression, class_name: &str, static_fields: &[String]) {
    match &mut expr.kind {
        ExprKind::Ident(n) => {
            if static_fields.iter().any(|f| f == n) {
                let name = n.clone();
                expr.kind = ExprKind::Member {
                    object: Box::new(if class_name == "__dart_instance" {
                        Expression::new(ExprKind::This)
                    } else {
                        Expression::ident(class_name)
                    }),
                    field: name,
                    null_safe: false,
                };
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_static_idents_expr(left, class_name, static_fields);
            rewrite_static_idents_expr(right, class_name, static_fields);
        }
        ExprKind::Unary { expr: inner, .. } => {
            rewrite_static_idents_expr(inner, class_name, static_fields)
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_static_idents_expr(callee, class_name, static_fields);
            for a in args.iter_mut() {
                rewrite_static_idents_expr(&mut a.value, class_name, static_fields);
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_static_idents_expr(object, class_name, static_fields)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_static_idents_expr(object, class_name, static_fields);
            rewrite_static_idents_expr(index, class_name, static_fields);
        }
        ExprKind::Assign { target, value } => {
            rewrite_static_idents_expr(target, class_name, static_fields);
            rewrite_static_idents_expr(value, class_name, static_fields);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_static_idents_expr(cond, class_name, static_fields);
            rewrite_static_idents_expr(then, class_name, static_fields);
            rewrite_static_idents_expr(else_, class_name, static_fields);
        }
        ExprKind::Array(elems) => {
            for e in elems.iter_mut() {
                rewrite_static_idents_expr(&mut e.value, class_name, static_fields);
            }
        }
        ExprKind::Object(props) => {
            for p in props.iter_mut() {
                if let vybe_ast::ObjectProperty::KeyValue { value, .. } = p {
                    rewrite_static_idents_expr(value, class_name, static_fields);
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(value) | InterpolPart::Formatted(value, _) = part {
                    rewrite_static_idents_expr(value, class_name, static_fields);
                }
            }
        }
        // Constructor ARGUMENTS only — the constructee is a type name, and
        // rewriting it would rebind a class that happens to share a field's
        // spelling. Missing this arm is why every operator overload that
        // builds its result (`operator +(o) => A(v + o.v)`, the shape all of
        // them use) read an unqualified `v` as a global: `return v` was
        // rewritten, `return A(v)` was not, so the body produced NaN.
        ExprKind::New { args, .. } => {
            for arg in args.iter_mut() {
                rewrite_static_idents_expr(&mut arg.value, class_name, static_fields);
            }
        }
        ExprKind::Await(inner) | ExprKind::Spread(inner) => {
            rewrite_static_idents_expr(inner, class_name, static_fields)
        }
        ExprKind::Tuple(items) | ExprKind::Sequence(items) => {
            for item in items.iter_mut() {
                rewrite_static_idents_expr(item, class_name, static_fields);
            }
        }
        // Descend into closures: a lambda inside an instance method (an event
        // handler, a `setState(() {...})` body) may reference sibling members
        // unqualified too. Inside a closure `this` is the DYNAMIC call-time
        // receiver — undefined when a host callback (GUI Click, forEach) fires
        // it — so route members through `_vybeSelf`, a local capturing `this`
        // at method entry that the closure captures as an upvalue.
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(value) => {
                rewrite_static_idents_expr(value, "_vybeSelf", static_fields)
            }
            LambdaBody::Block(stmts) => {
                for stmt in stmts.iter_mut() {
                    rewrite_static_idents(stmt, "_vybeSelf", static_fields);
                }
            }
        },
        _ => {}
    }
}

/// `let _vybeSelf = this;` — captures the instance so closures in a method
/// reach members through a real (upvalue-captured) local.
fn self_capture_stmt() -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("_vybeSelf".to_string()),
            type_hint: None,
            init: Some(Expression::new(ExprKind::This)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

fn stmts_contain_lambda(stmts: &[Statement]) -> bool {
    stmts.iter().any(stmt_contains_lambda)
}

fn stmt_contains_lambda(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(e) | StmtKind::Return(Some(e)) => expr_contains_lambda(e),
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .any(|d| d.init.as_ref().is_some_and(expr_contains_lambda)),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
            ..
        } => {
            expr_contains_lambda(cond)
                || stmts_contain_lambda(then_body)
                || elifs
                    .iter()
                    .any(|(c, b)| expr_contains_lambda(c) || stmts_contain_lambda(b))
                || else_body.as_ref().is_some_and(|b| stmts_contain_lambda(b))
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            init.as_deref().is_some_and(stmt_contains_lambda)
                || cond.as_ref().is_some_and(expr_contains_lambda)
                || update.as_ref().is_some_and(expr_contains_lambda)
                || stmts_contain_lambda(body)
        }
        StmtKind::ForIn { iter, body, .. } => {
            expr_contains_lambda(iter) || stmts_contain_lambda(body)
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            expr_contains_lambda(cond) || stmts_contain_lambda(body)
        }
        StmtKind::Block(stmts) => stmts_contain_lambda(stmts),
        _ => false,
    }
}

fn expr_contains_lambda(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Lambda { .. } => true,
        ExprKind::Binary { left, right, .. } => {
            expr_contains_lambda(left) || expr_contains_lambda(right)
        }
        ExprKind::Unary { expr, .. } => expr_contains_lambda(expr),
        ExprKind::Call { callee, args, .. } => {
            expr_contains_lambda(callee) || args.iter().any(|a| expr_contains_lambda(&a.value))
        }
        ExprKind::Member { object, .. } => expr_contains_lambda(object),
        ExprKind::Index { object, index, .. } => {
            expr_contains_lambda(object) || expr_contains_lambda(index)
        }
        ExprKind::Assign { target, value } => {
            expr_contains_lambda(target) || expr_contains_lambda(value)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            expr_contains_lambda(cond) || expr_contains_lambda(then) || expr_contains_lambda(else_)
        }
        ExprKind::Array(elems) => elems.iter().any(|e| expr_contains_lambda(&e.value)),
        _ => false,
    }
}

/// Dart permits unqualified reads, writes, and calls of instance members in
/// instance methods. The common ECMA class path receives `this` ambiently, so
/// make that receiver explicit in the AST instead of relying on a synthetic
/// positional self argument.
fn rewrite_instance_member_idents(members: &mut [ClassMember], extra_members: &[&str]) {
    let mut instance_members: Vec<String> =
        members.iter().filter_map(instance_member_name).collect();
    instance_members.extend(extra_members.iter().map(|name| (*name).to_string()));
    if instance_members.is_empty() {
        return;
    }
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl {
                    body, modifiers, ..
                } = &mut stmt.kind
                {
                    if !modifiers.is_static {
                        // A host-invoked closure (GUI Click, forEach, Future)
                        // clobbers the global `this`/`__js_this` and never
                        // restores it, so any `this.member` AFTER such a callback
                        // — and every member reference INSIDE it — reads garbage.
                        // When the method contains a closure, capture `this` into
                        // a real local (`_vybeSelf`) at entry and route EVERY
                        // member reference (inside and outside closures) through
                        // it, never relying on the fragile global receiver.
                        // Closure-free methods keep the plain `this.member` form.
                        if stmts_contain_lambda(body) {
                            body.insert(0, self_capture_stmt());
                            for stmt in body {
                                rewrite_static_idents(stmt, "_vybeSelf", &instance_members);
                            }
                        } else {
                            for stmt in body {
                                rewrite_static_idents(stmt, "__dart_instance", &instance_members);
                            }
                        }
                    }
                }
            }
            ClassMember::Property {
                name,
                getter: Some(body),
                modifiers,
                ..
            } if name == "hashCode" && !modifiers.is_static => {
                for stmt in body {
                    rewrite_static_idents(stmt, "__dart_instance", &instance_members);
                    rewrite_this_to_self_ident(stmt);
                }
            }
            ClassMember::Property {
                getter,
                setter,
                modifiers,
                ..
            } if !modifiers.is_static => {
                // A SETTER body needs the same rewrite as a getter's: `set v(n)
                // { _v = n; }` writes an instance field through a bare name.
                // Only the getter was covered, which went unnoticed while the
                // walker still copied mixin members into the using class — the
                // class-level pass caught the leftovers. It no longer copies.
                for stmt in getter.iter_mut().flatten().chain(
                    setter
                        .iter_mut()
                        .flat_map(|s: &mut PropertySetter| s.body.iter_mut()),
                ) {
                    rewrite_static_idents(stmt, "__dart_instance", &instance_members);
                    rewrite_this_to_self_ident(stmt);
                }
            }
            _ => {}
        }
    }
}

fn rewrite_this_to_self_ident(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            rewrite_this_to_self_ident_expr(expr);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_this_to_self_ident_expr(init);
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                rewrite_this_to_self_ident_expr(target);
            }
            rewrite_this_to_self_ident_expr(value);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_this_to_self_ident_expr(target);
            rewrite_this_to_self_ident_expr(value);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_this_to_self_ident_expr(cond);
            for stmt in then_body {
                rewrite_this_to_self_ident(stmt);
            }
            for (cond, body) in elifs {
                rewrite_this_to_self_ident_expr(cond);
                for stmt in body {
                    rewrite_this_to_self_ident(stmt);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_this_to_self_ident(stmt);
                }
            }
        }
        StmtKind::Block(body) => {
            for stmt in body {
                rewrite_this_to_self_ident(stmt);
            }
        }
        _ => {}
    }
}

fn rewrite_this_to_self_ident_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::This => {
            *expr = Expression::ident("this");
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_this_to_self_ident_expr(left);
            rewrite_this_to_self_ident_expr(right);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::YieldFrom(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Cast { expr, .. } => rewrite_this_to_self_ident_expr(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_this_to_self_ident_expr(cond);
            rewrite_this_to_self_ident_expr(then);
            rewrite_this_to_self_ident_expr(else_);
        }
        ExprKind::Member { object, .. } => rewrite_this_to_self_ident_expr(object),
        ExprKind::Index { object, index, .. } => {
            rewrite_this_to_self_ident_expr(object);
            rewrite_this_to_self_ident_expr(index);
        }
        ExprKind::Call { callee, args, .. }
        | ExprKind::New {
            class: callee,
            args,
        } => {
            rewrite_this_to_self_ident_expr(callee);
            for arg in args {
                rewrite_this_to_self_ident_expr(&mut arg.value);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_this_to_self_ident_expr(target);
            rewrite_this_to_self_ident_expr(value);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_this_to_self_ident_expr(&mut item.value);
                if let Some(key) = &mut item.key {
                    rewrite_this_to_self_ident_expr(key);
                }
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_this_to_self_ident_expr(key);
                        rewrite_this_to_self_ident_expr(value);
                    }
                    ObjectProperty::Spread(value) => rewrite_this_to_self_ident_expr(value),
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_this_to_self_ident_expr(item);
            }
        }
        _ => {}
    }
}

// ════════════════════════════════════════════════════════════════════════════
// noSuchMethod — Dart's missing-member hook
// ════════════════════════════════════════════════════════════════════════════

/// Which access shape produced an `Invocation`.
#[derive(Clone, Copy, PartialEq)]
enum NsmAccess {
    Method,
    Getter,
    Setter,
}

/// Dart sends a failed member access on a `dynamic` receiver to that object's
/// `noSuchMethod(Invocation)` instead of throwing. This lowers it the same way
/// PHP already lowers `__call` — `php/src/walker.rs::build_magic_call_rewrite`
/// — as a runtime test at the access site whose miss branch invokes the hook.
///
/// It is walker work, not shared-compiler work. The `CallMissing` protocol slot
/// that Dart, PHP and Ruby all register is a compile-time *role*; nothing in
/// the emitter reads it, and the lowering here is plain AST (a sequence, a
/// ternary, a `typeof`), so the shared compiler is untouched.
///
/// Inert unless some class in the program declares `noSuchMethod`: a Dart
/// program that never uses the feature comes out of this pass unchanged.
fn apply_no_such_method(body: &mut [Statement]) {
    if !nsm_module_declares_hook(body) {
        return;
    }
    let mut dynamic_vars: HashSet<String> = HashSet::new();
    nsm_rewrite_stmts(body, &mut dynamic_vars);
}

fn nsm_module_declares_hook(body: &[Statement]) -> bool {
    body.iter().any(|stmt| match &stmt.kind {
        StmtKind::ClassDecl { members, .. } => members.iter().any(|m| match m {
            ClassMember::Method(inner) => {
                matches!(&inner.kind, StmtKind::FunctionDecl { name, .. } if name == "noSuchMethod")
            }
            _ => false,
        }),
        _ => false,
    })
}

/// `#name` and a member's own spelling both render as Dart renders a `Symbol`:
/// `Symbol("name")`. Keeping them as that exact string is what makes
/// `inv.memberName == #doubleIt`, `inv.memberName.toString()` and
/// `inv.namedArguments[#mode]` all agree without a Symbol value type.
fn nsm_symbol(name: &str) -> String {
    format!("Symbol(\"{name}\")")
}

fn nsm_tmp(prefix: &str) -> String {
    use std::sync::atomic::{AtomicUsize, Ordering};
    static COUNTER: AtomicUsize = AtomicUsize::new(0);
    format!(
        "__dart_nsm_{prefix}{}",
        COUNTER.fetch_add(1, Ordering::Relaxed)
    )
}

/// A receiver is `dynamic` when it was declared so, or when it is the result of
/// a member access on something that is — Dart types both a missing-member call
/// and its result `dynamic`, which is what makes `c.next().end()` chain through
/// the hook and back out again.
fn nsm_is_dynamic(expr: &Expression, dynamic_vars: &HashSet<String>) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => dynamic_vars.contains(name),
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Member { object, .. } => nsm_is_dynamic(object, dynamic_vars),
            _ => false,
        },
        ExprKind::Member { object, .. } => nsm_is_dynamic(object, dynamic_vars),
        _ => false,
    }
}

/// The `Invocation` handed to the hook. Dart's `Invocation` is a class, but
/// every member the program can read off it is a plain field, so an object
/// literal carries it exactly — verified against the Dart SDK:
/// `Symbol("run") m=true g=false s=false pos=[1, 2] named={Symbol("mode"): fast}`.
fn nsm_invocation(member: &str, access: NsmAccess, args: &[Argument]) -> Expression {
    let mut positional = Vec::new();
    let mut named = Vec::new();
    for arg in args {
        match &arg.name {
            // Dart keys `namedArguments` by Symbol, and prints those keys as
            // `Symbol("mode")` — the same spelling `#mode` lowers to, so an
            // index by either finds the entry.
            Some(name) => named.push(ObjectProperty::KeyValue {
                key: Expression::string(&nsm_symbol(name)),
                value: arg.value.clone(),
            }),
            None => positional.push(ArrayElement {
                key: None,
                value: arg.value.clone(),
                spread: false,
                by_ref: false,
            }),
        }
    }
    // A setter's member name carries the `=`: real Dart reports `Symbol("value=")`
    // for `p.value = 1`, not `Symbol("value")`.
    let member_name = if access == NsmAccess::Setter {
        nsm_symbol(&format!("{member}="))
    } else {
        nsm_symbol(member)
    };
    let field = |name: &str, value: Expression| ObjectProperty::KeyValue {
        key: Expression::string(name),
        value,
    };
    Expression::new(ExprKind::Object(vec![
        field("memberName", Expression::string(&member_name)),
        field("isMethod", Expression::bool(access == NsmAccess::Method)),
        field("isGetter", Expression::bool(access == NsmAccess::Getter)),
        field("isSetter", Expression::bool(access == NsmAccess::Setter)),
        field("isAccessor", Expression::bool(access != NsmAccess::Method)),
        field(
            "positionalArguments",
            Expression::new(ExprKind::Array(positional)),
        ),
        field("namedArguments", Expression::new(ExprKind::Object(named))),
    ]))
}

fn nsm_member(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn nsm_hook_call(receiver: Expression, invocation: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(nsm_member(receiver, "noSuchMethod")),
        args: vec![Argument::positional(invocation)],
        optional: false,
    })
}

fn nsm_typeof_is(expr: Expression, type_name: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::StrictEq,
        left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(expr)))),
        right: Box::new(Expression::string(type_name)),
    })
}

/// `(t = obj, typeof t.m === "function" ? t.m(args) : t.noSuchMethod(inv))`
fn nsm_call_rewrite(object: Expression, field: &str, args: Vec<Argument>) -> Expression {
    let tmp = nsm_tmp("recv");
    let receiver = || Expression::ident(&tmp);
    let save = Expression::new(ExprKind::Assign {
        target: Box::new(receiver()),
        value: Box::new(object),
    });
    let direct = Expression::new(ExprKind::Call {
        callee: Box::new(nsm_member(receiver(), field)),
        args: args.clone(),
        optional: false,
    });
    let miss = nsm_hook_call(receiver(), nsm_invocation(field, NsmAccess::Method, &args));
    Expression::new(ExprKind::Sequence(vec![
        save,
        Expression::new(ExprKind::Ternary {
            cond: Box::new(nsm_typeof_is(nsm_member(receiver(), field), "function")),
            then: Box::new(direct),
            else_: Box::new(miss),
        }),
    ]))
}

/// `(t = obj, typeof t.p === "undefined" ? t.noSuchMethod(inv) : t.p)`
fn nsm_get_rewrite(object: Expression, field: &str) -> Expression {
    let tmp = nsm_tmp("recv");
    let receiver = || Expression::ident(&tmp);
    let save = Expression::new(ExprKind::Assign {
        target: Box::new(receiver()),
        value: Box::new(object),
    });
    let miss = nsm_hook_call(receiver(), nsm_invocation(field, NsmAccess::Getter, &[]));
    Expression::new(ExprKind::Sequence(vec![
        save,
        Expression::new(ExprKind::Ternary {
            cond: Box::new(nsm_typeof_is(nsm_member(receiver(), field), "undefined")),
            then: Box::new(miss),
            else_: Box::new(nsm_member(receiver(), field)),
        }),
    ]))
}

/// `(t = obj, v = val, typeof t.p === "undefined" ? t.noSuchMethod(inv) : (t.p = v))`
///
/// The value goes into its own temp so it is evaluated exactly once and in
/// source order, before the branch that decides where it lands.
fn nsm_set_rewrite(object: Expression, field: &str, value: Expression) -> Expression {
    let recv_tmp = nsm_tmp("recv");
    let value_tmp = nsm_tmp("val");
    let receiver = || Expression::ident(&recv_tmp);
    let held = || Expression::ident(&value_tmp);
    let save_receiver = Expression::new(ExprKind::Assign {
        target: Box::new(receiver()),
        value: Box::new(object),
    });
    let save_value = Expression::new(ExprKind::Assign {
        target: Box::new(held()),
        value: Box::new(value),
    });
    let miss = nsm_hook_call(
        receiver(),
        nsm_invocation(field, NsmAccess::Setter, &[Argument::positional(held())]),
    );
    let direct = Expression::new(ExprKind::Assign {
        target: Box::new(nsm_member(receiver(), field)),
        value: Box::new(held()),
    });
    Expression::new(ExprKind::Sequence(vec![
        save_receiver,
        save_value,
        Expression::new(ExprKind::Ternary {
            cond: Box::new(nsm_typeof_is(nsm_member(receiver(), field), "undefined")),
            then: Box::new(miss),
            else_: Box::new(direct),
        }),
    ]))
}

fn nsm_rewrite_stmts(stmts: &mut [Statement], dynamic_vars: &mut HashSet<String>) {
    for stmt in stmts.iter_mut() {
        nsm_rewrite_stmt(stmt, dynamic_vars);
    }
}

/// A nested body gets its own copy of the dynamic set: a `dynamic` local in one
/// function must not make a same-named local in another function dynamic.
fn nsm_rewrite_body(body: &mut [Statement], params: &[Param], dynamic_vars: &HashSet<String>) {
    let mut inner = dynamic_vars.clone();
    for param in params {
        if param.type_hint.as_deref() == Some("dynamic") {
            inner.insert(param.name.clone());
        }
    }
    nsm_rewrite_stmts(body, &mut inner);
}

fn nsm_rewrite_stmt(stmt: &mut Statement, dynamic_vars: &mut HashSet<String>) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => nsm_rewrite_expr(expr, dynamic_vars),
        StmtKind::Throw { expr, cause } => {
            for e in expr.iter_mut().chain(cause.iter_mut()) {
                nsm_rewrite_expr(e, dynamic_vars);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations.iter_mut() {
                if let Some(init) = &mut decl.init {
                    nsm_rewrite_expr(init, dynamic_vars);
                }
                if decl.type_hint.as_deref() == Some("dynamic") {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        // Recorded for the rest of the enclosing body. A later
                        // non-dynamic binding of the same name only costs a
                        // redundant runtime test — the direct branch still
                        // wins whenever the member exists.
                        dynamic_vars.insert(name.clone());
                    }
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            nsm_rewrite_expr(value, dynamic_vars);
            // `p.field = v` on a dynamic receiver is a setter invocation, and
            // the whole statement becomes the rewritten expression.
            if targets.len() == 1 {
                if let ExprKind::Member {
                    object,
                    field,
                    null_safe: false,
                } = &targets[0].kind
                {
                    if nsm_is_dynamic(object, dynamic_vars) {
                        let mut receiver = (**object).clone();
                        nsm_rewrite_expr(&mut receiver, dynamic_vars);
                        stmt.kind = StmtKind::Expr(nsm_set_rewrite(receiver, field, value.clone()));
                        return;
                    }
                }
            }
            for target in targets.iter_mut() {
                nsm_rewrite_expr(target, dynamic_vars);
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            nsm_rewrite_expr(target, dynamic_vars);
            nsm_rewrite_expr(value, dynamic_vars);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
            ..
        } => {
            nsm_rewrite_expr(cond, dynamic_vars);
            nsm_rewrite_stmts(then_body, dynamic_vars);
            for (elif_cond, body) in elifs.iter_mut() {
                nsm_rewrite_expr(elif_cond, dynamic_vars);
                nsm_rewrite_stmts(body, dynamic_vars);
            }
            if let Some(body) = else_body {
                nsm_rewrite_stmts(body, dynamic_vars);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            if let Some(init) = init.as_deref_mut() {
                nsm_rewrite_stmt(init, dynamic_vars);
            }
            if let Some(cond) = cond {
                nsm_rewrite_expr(cond, dynamic_vars);
            }
            if let Some(update) = update {
                nsm_rewrite_expr(update, dynamic_vars);
            }
            nsm_rewrite_stmts(body, dynamic_vars);
        }
        StmtKind::ForIn { iter, body, .. } => {
            nsm_rewrite_expr(iter, dynamic_vars);
            nsm_rewrite_stmts(body, dynamic_vars);
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            nsm_rewrite_expr(cond, dynamic_vars);
            nsm_rewrite_stmts(body, dynamic_vars);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            nsm_rewrite_expr(expr, dynamic_vars);
            for case in cases.iter_mut() {
                nsm_rewrite_stmts(&mut case.body, dynamic_vars);
            }
            if let Some(body) = default {
                nsm_rewrite_stmts(body, dynamic_vars);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            nsm_rewrite_stmts(body, dynamic_vars);
            for catch in catches.iter_mut() {
                nsm_rewrite_stmts(&mut catch.body, dynamic_vars);
            }
            for extra in else_body.iter_mut().chain(finally.iter_mut()) {
                nsm_rewrite_stmts(extra, dynamic_vars);
            }
        }
        StmtKind::Block(body) => nsm_rewrite_stmts(body, dynamic_vars),
        StmtKind::FunctionDecl { params, body, .. } => nsm_rewrite_body(body, params, dynamic_vars),
        StmtKind::ClassDecl { members, .. } => {
            for member in members.iter_mut() {
                match member {
                    ClassMember::Method(inner) => nsm_rewrite_stmt(inner, dynamic_vars),
                    ClassMember::Constructor { params, body, .. } => {
                        nsm_rewrite_body(body, params, dynamic_vars)
                    }
                    ClassMember::Property { getter, setter, .. } => {
                        if let Some(body) = getter {
                            nsm_rewrite_body(body, &[], dynamic_vars);
                        }
                        if let Some(setter) = setter {
                            nsm_rewrite_body(&mut setter.body, &[], dynamic_vars);
                        }
                    }
                    ClassMember::Field { init: Some(e), .. } => nsm_rewrite_expr(e, dynamic_vars),
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

fn nsm_rewrite_expr(expr: &mut Expression, dynamic_vars: &HashSet<String>) {
    // A call on a dynamic receiver is decided BEFORE descending into the
    // callee — otherwise the `Member` inside `c.next()` would be rewritten as
    // a getter first and the call would never be seen.
    if let ExprKind::Call { callee, args, .. } = &expr.kind {
        if let ExprKind::Member {
            object,
            field,
            null_safe: false,
        } = &callee.kind
        {
            if nsm_is_dynamic(object, dynamic_vars) {
                let mut receiver = (**object).clone();
                let field = field.clone();
                let mut args = args.clone();
                nsm_rewrite_expr(&mut receiver, dynamic_vars);
                for arg in args.iter_mut() {
                    nsm_rewrite_expr(&mut arg.value, dynamic_vars);
                }
                *expr = nsm_call_rewrite(receiver, &field, args);
                return;
            }
        }
    }
    if let ExprKind::Assign { target, value } = &expr.kind {
        if let ExprKind::Member {
            object,
            field,
            null_safe: false,
        } = &target.kind
        {
            if nsm_is_dynamic(object, dynamic_vars) {
                let mut receiver = (**object).clone();
                let field = field.clone();
                let mut held = (**value).clone();
                nsm_rewrite_expr(&mut receiver, dynamic_vars);
                nsm_rewrite_expr(&mut held, dynamic_vars);
                *expr = nsm_set_rewrite(receiver, &field, held);
                return;
            }
        }
    }
    if let ExprKind::Member {
        object,
        field,
        null_safe: false,
    } = &expr.kind
    {
        if nsm_is_dynamic(object, dynamic_vars) {
            let mut receiver = (**object).clone();
            let field = field.clone();
            nsm_rewrite_expr(&mut receiver, dynamic_vars);
            *expr = nsm_get_rewrite(receiver, &field);
            return;
        }
    }

    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            nsm_rewrite_expr(left, dynamic_vars);
            nsm_rewrite_expr(right, dynamic_vars);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::TypeOf(inner)
        | ExprKind::Await(inner)
        | ExprKind::Spread(inner) => nsm_rewrite_expr(inner, dynamic_vars),
        ExprKind::Call { callee, args, .. } => {
            nsm_rewrite_expr(callee, dynamic_vars);
            for arg in args.iter_mut() {
                nsm_rewrite_expr(&mut arg.value, dynamic_vars);
            }
        }
        ExprKind::New { class, args } => {
            nsm_rewrite_expr(class, dynamic_vars);
            for arg in args.iter_mut() {
                nsm_rewrite_expr(&mut arg.value, dynamic_vars);
            }
        }
        ExprKind::Member { object, .. } => nsm_rewrite_expr(object, dynamic_vars),
        ExprKind::Index { object, index, .. } => {
            nsm_rewrite_expr(object, dynamic_vars);
            nsm_rewrite_expr(index, dynamic_vars);
        }
        ExprKind::Assign { target, value } => {
            nsm_rewrite_expr(target, dynamic_vars);
            nsm_rewrite_expr(value, dynamic_vars);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            nsm_rewrite_expr(cond, dynamic_vars);
            nsm_rewrite_expr(then, dynamic_vars);
            nsm_rewrite_expr(else_, dynamic_vars);
        }
        ExprKind::Array(elements) => {
            for element in elements.iter_mut() {
                nsm_rewrite_expr(&mut element.value, dynamic_vars);
            }
        }
        ExprKind::Object(props) => {
            for prop in props.iter_mut() {
                match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        nsm_rewrite_expr(key, dynamic_vars);
                        nsm_rewrite_expr(value, dynamic_vars);
                    }
                    ObjectProperty::Spread(value) => nsm_rewrite_expr(value, dynamic_vars),
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Sequence(items) => {
            for item in items.iter_mut() {
                nsm_rewrite_expr(item, dynamic_vars);
            }
        }
        ExprKind::Lambda { body, params, .. } => {
            let mut inner = dynamic_vars.clone();
            for param in params.iter() {
                if param.type_hint.as_deref() == Some("dynamic") {
                    inner.insert(param.name.clone());
                }
            }
            match body {
                LambdaBody::Expr(e) => nsm_rewrite_expr(e, &inner),
                LambdaBody::Block(stmts) => nsm_rewrite_stmts(stmts, &mut inner),
            }
        }
        _ => {}
    }
}

/// The value half of a `const` expression, built with no canonicalization —
/// the caller has already decided that THIS occurrence is the one that gets
/// built, and every other occurrence becomes a reference to it.
thread_local! {
    /// Canonicalized `const` expressions for the program being parsed, in
    /// creation order: `(source key, lowered value, binding name)`.
    static DART_CONST_POOL: std::cell::RefCell<Vec<(String, Expression, String)>> =
        const { std::cell::RefCell::new(Vec::new()) };
}

/// Collapse whitespace so `const [1,2]` and `const [1, 2]` are one constant.
fn dart_const_key(src: &str) -> String {
    src.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn dart_const_pool_lookup(src: &str) -> Option<String> {
    let key = dart_const_key(src);
    DART_CONST_POOL.with(|pool| {
        pool.borrow()
            .iter()
            .find(|(k, _, _)| *k == key)
            .map(|(_, _, name)| name.clone())
    })
}

fn dart_const_pool_insert(src: String, value: Expression) -> String {
    let key = dart_const_key(&src);
    DART_CONST_POOL.with(|pool| {
        let mut pool = pool.borrow_mut();
        let name = format!("__dart_const_{}", pool.len());
        pool.push((key, value, name.clone()));
        name
    })
}

/// `var __dart_const_N = <value>;` for each canonicalized const, in creation
/// order — an inner const is pooled before the outer one that contains it, so
/// creation order is already dependency order.
fn dart_const_pool_declarations() -> Vec<Statement> {
    DART_CONST_POOL.with(|pool| {
        pool.borrow()
            .iter()
            .map(|(_, value, name)| {
                Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(name.clone()),
                        type_hint: None,
                        init: Some(value.clone()),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                })
            })
            .collect()
    })
}

fn walk_const_expression_value(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // const ClassName(args) — treat same as new
    let mut class_parts: Vec<String> = Vec::new();
    let mut args = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            // `const [...]` / `const {...}` — the collection literal IS the
            // value; `const` only decides that it is canonicalized.
            Rule::list_literal | Rule::map_or_set_literal => {
                return walk_expr_kind(p);
            }
            Rule::ident_name => class_parts.push(p.as_str().to_string()),
            Rule::type_args => {}
            Rule::argument_list => args = walk_arguments(p)?,
            Rule::const_kw => {}
            _ => {}
        }
    }
    let class_name = class_parts.join(".");
    if let Some((ty, ctor)) = class_name.split_once('.') {
        if let Some(kind) = dart_flutter_named_ctor(ty, ctor, &args) {
            return Ok(kind);
        }
    }
    inject_flutter_defaults(&class_name, &mut args);
    Ok(ExprKind::New {
        class: Box::new(Expression::ident(&class_name)),
        args,
    })
}

fn walk_mixin_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::type_params => consume_dart_type_params(p),
            Rule::on_clause => {
                for ta in p.into_inner() {
                    if ta.as_rule() == Rule::type_annotation_list {
                        for t in ta.into_inner() {
                            if t.as_rule() == Rule::type_annotation {
                                parents.push(extract_type_name(&t));
                            }
                        }
                    }
                }
            }
            Rule::implements_clause => {
                for ta in p.into_inner() {
                    if ta.as_rule() == Rule::type_annotation_list {
                        for t in ta.into_inner() {
                            if t.as_rule() == Rule::type_annotation {
                                interfaces.push(extract_type_name(&t));
                            }
                        }
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    match m.as_rule() {
                        Rule::constructor_declaration
                        | Rule::operator_declaration
                        | Rule::getter_declaration
                        | Rule::setter_declaration
                        | Rule::method_declaration
                        | Rule::field_declaration => {
                            if let Some(member) = walk_class_member(m, &name)? {
                                members.push(member);
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    rewrite_instance_member_idents(&mut members, &[]);

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn walk_extension_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // Extension on Type { members } — treat as a class with the target type as parent
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::type_annotation => {
                parents.push(extract_type_name(&p));
            }
            Rule::constructor_declaration
            | Rule::operator_declaration
            | Rule::getter_declaration
            | Rule::setter_declaration
            | Rule::method_declaration
            | Rule::field_declaration => {
                if let Some(member) = walk_class_member(p, &name)? {
                    members.push(member);
                }
            }
            _ => {}
        }
    }

    if name.is_empty() {
        name = "__extension__".to_string();
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers {
            is_static: true,
            ..Default::default()
        },
        decorators: vec![],
    })
}

fn walk_extension_type_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut representation_name = String::new();
    let mut representation_type: Option<String> = None;
    let mut interfaces = Vec::new();
    let mut members = Vec::new();

    for part in pair.into_inner() {
        match part.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = part.as_str().to_string();
                } else if representation_name.is_empty() {
                    representation_name = part.as_str().to_string();
                }
            }
            Rule::type_annotation => {
                if representation_type.is_none() {
                    representation_type = Some(extract_type_name(&part));
                }
            }
            Rule::implements_clause => {
                for item in part.into_inner() {
                    if item.as_rule() == Rule::type_annotation_list {
                        for ty in item.into_inner() {
                            if ty.as_rule() == Rule::type_annotation {
                                interfaces.push(extract_type_name(&ty));
                            }
                        }
                    }
                }
            }
            Rule::class_body => {
                for member in part.into_inner() {
                    match member.as_rule() {
                        Rule::constructor_declaration
                        | Rule::operator_declaration
                        | Rule::getter_declaration
                        | Rule::setter_declaration
                        | Rule::method_declaration
                        | Rule::field_declaration => {
                            if let Some(member) = walk_class_member(member, &name)? {
                                members.push(member);
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    if representation_name.is_empty() {
        representation_name = "representation".to_string();
    }
    rewrite_instance_member_idents(&mut members, &[representation_name.as_str()]);
    let type_hint = representation_type.unwrap_or_else(|| "dynamic".to_string());
    members.insert(
        0,
        ClassMember::Field {
            name: representation_name.clone(),
            type_hint: Some(type_hint.clone()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
        },
    );
    members.insert(
        1,
        ClassMember::Constructor {
            name: None,
            params: vec![Param {
                name: representation_name.clone(),
                type_hint: Some(type_hint.into()),
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: vec![Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Assign {
                    target: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: representation_name.clone(),
                        null_safe: false,
                    })),
                    value: Box::new(Expression::ident(&representation_name)),
                },
            )))],
            base_args: None,
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        },
    );
    Ok(StmtKind::ClassDecl {
        name,
        parents: Vec::new(),
        interfaces,
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn walk_enum_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut members: Vec<EnumMember> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut body_members: Vec<ClassMember> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::enum_values => {
                for vp in p.into_inner() {
                    match vp.as_rule() {
                        Rule::ident_name => {
                            members.push(EnumMember {
                                name: vp.as_str().to_string(),
                                value: None,
                                constructor_args: Vec::new(),
                            });
                        }
                        Rule::enum_value => {
                            let mut value_name = String::new();
                            let mut constructor_args = Vec::new();
                            for inner in vp.into_inner() {
                                match inner.as_rule() {
                                    Rule::ident_name if value_name.is_empty() => {
                                        value_name = inner.as_str().to_string();
                                    }
                                    Rule::argument_list => {
                                        constructor_args = walk_arguments(inner)?
                                            .into_iter()
                                            .map(|arg| arg.value)
                                            .collect();
                                    }
                                    _ => {}
                                }
                            }
                            members.push(EnumMember {
                                name: value_name,
                                value: None,
                                constructor_args,
                            });
                        }
                        _ => {}
                    }
                }
            }
            Rule::enum_clauses => {
                let raw = p.as_str();
                if let Some(idx) = raw.find("implements") {
                    let tail = &raw[idx + "implements".len()..];
                    interfaces.extend(
                        tail.split(',')
                            .map(str::trim)
                            .filter(|name| !name.is_empty())
                            .map(|name| name.to_string()),
                    );
                }
            }
            Rule::class_member
            | Rule::constructor_declaration
            | Rule::operator_declaration
            | Rule::getter_declaration
            | Rule::setter_declaration
            | Rule::method_declaration
            | Rule::field_declaration => {
                if let Some(member) = walk_class_member(p, &name)? {
                    body_members.push(member);
                }
            }
            Rule::type_params => consume_dart_type_params(p),
            _ => {}
        }
    }

    // Ordinary enums use the compiler's compact Dart enum representation.
    // Enhanced enums need real instances: their value constructors initialize
    // fields and their methods run with an instance receiver. Normalize those
    // to the shared class model, as PHP does for its enum singletons.
    let is_enhanced = !body_members.is_empty()
        || members
            .iter()
            .any(|member| !member.constructor_args.is_empty());
    if !is_enhanced {
        return Ok(StmtKind::EnumDecl {
            name,
            interfaces,
            members,
            visibility: Visibility::Public,
            is_flags: false,
            backing_type: None,
            body_members,
            decorators: vec![],
        });
    }

    rewrite_instance_member_idents(&mut body_members, &["index", "name"]);

    let static_modifiers = Modifiers {
        is_static: true,
        ..Modifiers::default()
    };
    let assign_this = |field: &str, value: Expression| {
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: field.to_string(),
                null_safe: false,
            })),
            value: Box::new(value),
        })))
    };
    let enum_param = |param_name: &str| Param {
        name: param_name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };

    let mut class_members = body_members;
    class_members.insert(
        0,
        ClassMember::Field {
            name: "index".to_string(),
            type_hint: Some("int".to_string()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
        },
    );
    class_members.insert(
        1,
        ClassMember::Field {
            name: "name".to_string(),
            type_hint: Some("String".to_string()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
        },
    );

    let constructor_index = class_members
        .iter()
        .position(|member| matches!(member, ClassMember::Constructor { name: None, .. }));
    if let Some(index) = constructor_index {
        if let ClassMember::Constructor { params, body, .. } = &mut class_members[index] {
            params.push(enum_param("__enum_index"));
            params.push(enum_param("__enum_name"));
            body.insert(0, assign_this("name", Expression::ident("__enum_name")));
            body.insert(0, assign_this("index", Expression::ident("__enum_index")));
        }
    } else {
        class_members.insert(
            2,
            ClassMember::Constructor {
                name: None,
                params: vec![enum_param("__enum_index"), enum_param("__enum_name")],
                body: vec![
                    assign_this("index", Expression::ident("__enum_index")),
                    assign_this("name", Expression::ident("__enum_name")),
                ],
                base_args: None,
                initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
                visibility: Visibility::Public,
            },
        );
    }

    let mut values = Vec::new();
    for (index, member) in members.iter().enumerate() {
        let mut args: Vec<Argument> = member
            .constructor_args
            .iter()
            .cloned()
            .map(Argument::positional)
            .collect();
        args.push(Argument::positional(Expression::new(ExprKind::Lit(
            Literal::Int(index as i64),
        ))));
        args.push(Argument::positional(Expression::string(&member.name)));
        let singleton = Expression::new(ExprKind::New {
            class: Box::new(Expression::ident(&name)),
            args,
        });
        class_members.push(ClassMember::Field {
            name: member.name.clone(),
            type_hint: Some(name.clone()),
            init: Some(singleton),
            modifiers: static_modifiers.clone(),
            with_events: false,
            array_bounds: None,
        });
        values.push(ArrayElement {
            key: None,
            value: Expression::new(ExprKind::StaticAccess {
                class: Box::new(Expression::ident(&name)),
                member: Box::new(Expression::ident(&member.name)),
            }),
            spread: false,
            by_ref: false,
        });
    }
    class_members.push(ClassMember::Field {
        name: "values".to_string(),
        type_hint: None,
        init: Some(Expression::new(ExprKind::Array(values))),
        modifiers: static_modifiers,
        with_events: false,
        array_bounds: None,
    });

    Ok(StmtKind::ClassDecl {
        name,
        // `Enum` is metadata in Dart, not a constructible runtime parent.
        // Adding it here sends the shared class emitter down its derived
        // constructor path and drops the receiver of bound instance methods.
        parents: Vec::new(),
        interfaces,
        members: class_members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Class members
// ════════════════════════════════════════════════════════════════════════════

fn walk_class_member(pair: Pair<Rule>, class_name: &str) -> Result<Option<ClassMember>, String> {
    match pair.as_rule() {
        Rule::constructor_declaration => Ok(Some(walk_constructor(pair, class_name)?)),
        Rule::method_declaration => Ok(Some(walk_method(pair)?)),
        Rule::field_declaration => walk_field(pair),
        Rule::getter_declaration => Ok(Some(walk_getter(pair)?)),
        Rule::setter_declaration => Ok(Some(walk_setter(pair)?)),
        Rule::operator_declaration => Ok(Some(walk_operator(pair)?)),
        Rule::annotation => Ok(None),
        _ => Ok(None),
    }
}

fn walk_member_modifiers(pair: &Pair<Rule>) -> Modifiers {
    let txt = pair.as_str();
    let mut m = Modifiers::default();
    if txt.contains("static") {
        m.is_static = true;
    }
    if txt.contains("abstract") {
        m.is_abstract = true;
    }
    if txt.contains("override") {
        m.is_override = true;
    }
    if txt.contains("final") {
        m.is_readonly = true;
    }
    m
}

fn walk_constructor(pair: Pair<Rule>, class_name: &str) -> Result<ClassMember, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut base_args: Option<Vec<Expression>> = None;
    let mut this_params: Vec<String> = Vec::new();
    let mut super_params: Vec<String> = Vec::new();
    let mut field_inits: Vec<Statement> = Vec::new();
    let mut is_factory = false;
    let mut _named_ctor: Option<String> = None;
    let mut found_name = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => {
                let txt = p.as_str();
                if txt.contains("factory") {
                    is_factory = true;
                }
            }
            Rule::const_kw => {}
            Rule::factory_kw => is_factory = true,
            Rule::ident_name => {
                if !found_name {
                    found_name = true;
                    // First ident is the class name (or named ctor prefix)
                } else {
                    // Second ident is the named constructor suffix
                    _named_ctor = Some(p.as_str().to_string());
                }
            }
            Rule::param_list => {
                for pg in p.into_inner() {
                    match pg.as_rule() {
                        Rule::param_group => {
                            for inner in pg.into_inner() {
                                match inner.as_rule() {
                                    Rule::param => {
                                        let (param, is_this, is_super) =
                                            walk_param_with_this(inner)?;
                                        if is_this {
                                            this_params.push(param.name.clone());
                                        }
                                        if is_super {
                                            super_params.push(param.name.clone());
                                        }
                                        params.push(param);
                                    }
                                    Rule::optional_positional_params | Rule::named_params => {
                                        for op in inner.into_inner() {
                                            if op.as_rule() == Rule::param {
                                                let (mut param, is_this, is_super) =
                                                    walk_param_with_this(op)?;
                                                param.is_optional = true;
                                                if is_this {
                                                    this_params.push(param.name.clone());
                                                }
                                                if is_super {
                                                    super_params.push(param.name.clone());
                                                }
                                                params.push(param);
                                            }
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                        Rule::param => {
                            let (param, is_this, is_super) = walk_param_with_this(pg)?;
                            if is_this {
                                this_params.push(param.name.clone());
                            }
                            if is_super {
                                super_params.push(param.name.clone());
                            }
                            params.push(param);
                        }
                        _ => {}
                    }
                }
            }
            Rule::initializer_list => {
                for init in p.into_inner() {
                    if init.as_rule() == Rule::initializer {
                        let inner = init.into_inner().next();
                        if let Some(ini) = inner {
                            match ini.as_rule() {
                                Rule::super_call_initializer => {
                                    let mut args = Vec::new();
                                    for sp in ini.into_inner() {
                                        if sp.as_rule() == Rule::argument_list {
                                            args = walk_arguments(sp)?
                                                .into_iter()
                                                .map(|a| a.value)
                                                .collect();
                                        }
                                    }
                                    base_args = Some(args);
                                }
                                Rule::this_redirect_initializer => {
                                    // `Point.origin() : this(0, 0)` — named
                                    // constructor redirecting to the unnamed
                                    // (or another named) constructor. Walker
                                    // lowers to: `var _self = ClassName(args);
                                    // return _self;` so the named ctor becomes
                                    // a factory-style static method.
                                    let mut redirect_target = None;
                                    let mut redirect_args = Vec::new();
                                    for sp in ini.into_inner() {
                                        match sp.as_rule() {
                                            Rule::ident_name => {
                                                redirect_target = Some(sp.as_str().to_string())
                                            }
                                            Rule::argument_list => {
                                                redirect_args = walk_arguments(sp)?
                                            }
                                            _ => {}
                                        }
                                    }
                                    let new_class = match redirect_target {
                                        Some(name) => Expression::new(ExprKind::Member {
                                            object: Box::new(Expression::ident(class_name)),
                                            field: name,
                                            null_safe: false,
                                        }),
                                        None => Expression::ident(class_name),
                                    };
                                    field_inits.push(Statement::new(StmtKind::Return(Some(
                                        Expression::new(ExprKind::New {
                                            class: Box::new(new_class),
                                            args: redirect_args,
                                        }),
                                    ))));
                                    is_factory = true;
                                }
                                Rule::field_initializer => {
                                    let mut field_name = String::new();
                                    let mut value_expr = None;
                                    for fp in ini.into_inner() {
                                        match fp.as_rule() {
                                            Rule::ident_name => {
                                                field_name = fp.as_str().to_string()
                                            }
                                            Rule::assignment_expression => {
                                                value_expr = Some(walk_expression(fp)?);
                                            }
                                            _ => {}
                                        }
                                    }
                                    if let Some(val) = value_expr {
                                        // Synthesize: this.field = expr;
                                        field_inits.push(Statement::new(StmtKind::Expr(
                                            Expression::new(ExprKind::Assign {
                                                target: Box::new(Expression::new(
                                                    ExprKind::Member {
                                                        object: Box::new(Expression::new(
                                                            ExprKind::This,
                                                        )),
                                                        field: field_name,
                                                        null_safe: false,
                                                    },
                                                )),
                                                value: Box::new(val),
                                            }),
                                        )));
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
            Rule::function_body_block => {
                body = walk_statement_into_body(p)?;
            }
            _ => {}
        }
    }

    if base_args.is_none() && !super_params.is_empty() {
        base_args = Some(
            super_params
                .iter()
                .map(|name| Expression::ident(name))
                .collect(),
        );
    }

    // Synthesize this.field = field assignments for this.* params
    let mut this_assigns: Vec<Statement> = this_params
        .iter()
        .map(|name| {
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: name.clone(),
                    null_safe: false,
                })),
                value: Box::new(Expression::ident(name)),
            })))
        })
        .collect();

    // Prepend: this.field assignments, then initializer list assignments, then body
    let mut full_body = Vec::new();
    full_body.append(&mut this_assigns);
    full_body.append(&mut field_inits);
    full_body.append(&mut body);

    if let (true, Some(named)) = (is_factory, _named_ctor.clone()) {
        // `factory Box.empty()` — reached as `Box.empty(...)`, so a static
        // method returning an instance is exactly its shape.
        Ok(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: named,
                params,
                return_type: Some(class_name.to_string()),
                body: full_body,
                modifiers: Modifiers {
                    is_static: true,
                    ..Default::default()
                },
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: false,
            },
        ))))
    } else if is_factory {
        // `factory Box(...)` — an UNNAMED factory *is* what `Box(...)` runs,
        // so it has to be the constructor. Its body always returns the
        // instance, and an explicit `return` from a constructor body yields
        // that value instead of the freshly allocated `this` — which is the
        // whole point of a factory (caches, singletons, subtype dispatch).
        // Naming it as a static method left it unreachable.
        Ok(ClassMember::Constructor {
            name: None,
            params,
            body: full_body,
            base_args,
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        })
    } else {
        // `Point.origin()` — a named generative constructor. Carrying the
        // name (rather than dropping it) is what lets the class keep both it
        // and the unnamed ctor: they are different constructors, not
        // overloads, and both are arity-0 here.
        Ok(ClassMember::Constructor {
            name: _named_ctor,
            params,
            body: full_body,
            base_args,
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        })
    }
}

fn walk_method(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut return_type: Option<String> = None;
    let mut modifiers = Modifiers::default();
    let mut is_async = false;
    let mut is_generator = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => modifiers = walk_member_modifiers(&p),
            Rule::type_annotation => {
                if return_type.is_none() {
                    return_type = Some(extract_type_name(&p));
                }
            }
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                }
            }
            Rule::type_params => consume_dart_type_params(p),
            Rule::param_list => params = walk_params(p)?,
            Rule::async_kw => is_async = true,
            Rule::generator_marker => is_generator = true,
            Rule::function_body => {
                let is_abstract_body = p.as_str().trim() == ";";
                body = walk_function_body(p)?;
                if is_abstract_body {
                    modifiers.is_abstract = true;
                }
            }
            _ => {}
        }
    }

    if is_async && !is_generator && body.is_empty() {
        body.push(Statement::new(StmtKind::Return(Some(dart_future_value(
            Expression::null(),
        )))));
    }
    is_generator = is_generator || body_has_yield(&body);

    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            handles: Vec::new(),
            is_async,
            is_generator,
            is_sub: false,
        },
    ))))
}

fn walk_field(pair: Pair<Rule>) -> Result<Option<ClassMember>, String> {
    // field_declaration = {
    //     member_modifiers ~ type_annotation? ~ ident_name ~ ("=" ~ assignment_expression)?
    //     ~ ("," ~ ident_name ~ ("=" ~ assignment_expression)?)* ~ ";"
    // }
    let mut modifiers = Modifiers::default();
    let mut type_hint: Option<String> = None;
    let mut fields: Vec<(String, Option<Expression>)> = Vec::new();
    let mut current_name = String::new();
    let mut is_const = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => {
                modifiers = walk_member_modifiers(&p);
                let txt = p.as_str();
                if txt.contains("const") {
                    is_const = true;
                }
            }
            Rule::type_annotation => {
                if type_hint.is_none() {
                    type_hint = Some(extract_type_name(&p));
                }
            }
            Rule::ident_name => {
                if !current_name.is_empty() {
                    fields.push((current_name.clone(), None));
                }
                current_name = p.as_str().to_string();
            }
            Rule::assignment_expression => {
                let init = walk_expression(p)?;
                fields.push((current_name.clone(), Some(init)));
                current_name = String::new();
            }
            _ => {}
        }
    }
    if !current_name.is_empty() {
        fields.push((current_name, None));
    }

    // Return first field (most common case: single field)
    // For multiple fields in one declaration, we return the first and
    // the caller should handle multi-field, but our grammar walks them individually.
    if let Some((name, init)) = fields.into_iter().next() {
        if is_const {
            Ok(Some(ClassMember::Const {
                name,
                type_hint: type_hint.clone(),
                value: init.unwrap_or(Expression::null()),
                visibility: modifiers.visibility,
            }))
        } else {
            Ok(Some(ClassMember::Field {
                name,
                type_hint,
                init,
                modifiers,
                with_events: false,
                array_bounds: None,
            }))
        }
    } else {
        Ok(None)
    }
}

fn walk_getter(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut body = Vec::new();
    let mut modifiers = Modifiers::default();
    let mut return_type: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => modifiers = walk_member_modifiers(&p),
            Rule::type_annotation => return_type = Some(extract_type_name(&p)),
            Rule::get_keyword => {}
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::async_kw => {}
            Rule::function_body => body = walk_function_body(p)?,
            _ => {}
        }
    }

    Ok(ClassMember::Property {
        name,
        type_hint: return_type,
        getter: Some(body),
        setter: None,
        is_auto: false,
        modifiers,
    })
}

fn walk_setter(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut modifiers = Modifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => modifiers = walk_member_modifiers(&p),
            Rule::type_annotation => {}
            Rule::set_keyword => {}
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::async_kw => {}
            Rule::function_body => body = walk_function_body(p)?,
            _ => {}
        }
    }

    let param = if let Some(p) = params.into_iter().next() {
        p
    } else {
        Param {
            name: "value".into(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }
    };

    Ok(ClassMember::Property {
        name,
        type_hint: None,
        getter: None,
        setter: Some(PropertySetter { param, body }),
        is_auto: false,
        modifiers,
    })
}

fn walk_operator(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut op_name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut modifiers = Modifiers::default();
    let mut return_type: Option<String> = None;
    let mut is_async = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::member_modifiers => modifiers = walk_member_modifiers(&p),
            Rule::type_annotation => {
                if return_type.is_none() {
                    return_type = Some(extract_type_name(&p));
                }
            }
            Rule::operator_symbol => {
                op_name = match p.as_str().trim() {
                    "+" => "operator+".to_string(),
                    "-" => "operator-".to_string(),
                    "*" => "operator*".to_string(),
                    "/" => "operator/".to_string(),
                    "~/" => "operator~/".to_string(),
                    "%" => "operator%".to_string(),
                    "==" => "__eq__".to_string(),
                    "!=" => "operator!=".to_string(),
                    "<" => "operator<".to_string(),
                    ">" => "operator>".to_string(),
                    "<=" => "operator<=".to_string(),
                    ">=" => "operator>=".to_string(),
                    "[]" => "__getitem__".to_string(),
                    "[]=" => "__setitem__".to_string(),
                    "~" => "operator~".to_string(),
                    "&" => "operator&".to_string(),
                    "|" => "operator|".to_string(),
                    "^" => "operator^".to_string(),
                    "<<" => "operator<<".to_string(),
                    ">>" => "operator>>".to_string(),
                    ">>>" => "operator>>>".to_string(),
                    other => format!("operator{}", other),
                };
            }
            Rule::param_list => params = walk_params(p)?,
            Rule::async_kw => is_async = true,
            Rule::function_body => body = walk_function_body(p)?,
            _ => {}
        }
    }

    // Dart spells unary minus and binary minus with the same token and tells
    // them apart by arity: `operator -()` negates, `operator -(other)`
    // subtracts. They are different methods, so a zero-parameter `-` is
    // `__neg__` — otherwise a class defining both binds only one of them.
    if op_name == "operator-" && params.is_empty() {
        op_name = "operator-@unary".to_string();
    }

    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: op_name,
            params,
            return_type,
            body,
            modifiers,
            handles: Vec::new(),
            is_async,
            is_generator: false,
            is_sub: false,
        },
    ))))
}

// ════════════════════════════════════════════════════════════════════════════
// Control flow
// ════════════════════════════════════════════════════════════════════════════

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond_pair = inner.next().ok_or("if: missing cond")?;
    let mut case_bindings = HashMap::new();
    let cond = if cond_pair.as_rule() == Rule::if_case_condition {
        let mut parts = cond_pair.into_inner();
        let subject = walk_expression(parts.next().ok_or("if-case: missing subject")?)?;
        let mut analysis = analyze_dart_pattern(
            parts
                .find(|p| p.as_rule() == Rule::pattern)
                .ok_or("if-case: missing pattern")?,
            &subject,
        )?;
        if let Some(guard_pair) = parts
            .find(|p| p.as_rule() == Rule::when_guard)
            .and_then(|p| {
                p.into_inner()
                    .find(|c| c.as_rule() == Rule::conditional_expression)
            })
        {
            let guard =
                substitute_pattern_bindings(walk_expression(guard_pair)?, &analysis.bindings);
            analysis.cond = and_expr(analysis.cond, guard);
        }
        case_bindings = analysis.bindings;
        analysis.cond
    } else {
        walk_expression(cond_pair)?
    };
    let then_stmt = inner.next().ok_or("if: missing body")?;
    let mut then_body = walk_statement_into_body(then_stmt)?;
    if !case_bindings.is_empty() {
        for stmt in then_body.iter_mut() {
            substitute_pattern_bindings_stmt(stmt, &case_bindings);
        }
    }

    // else clause
    let else_body = match inner.next() {
        Some(else_pair) => {
            if else_pair.as_rule() == Rule::else_clause {
                let else_stmt = else_pair.into_inner().next().ok_or("else: missing body")?;
                Some(walk_statement_into_body(else_stmt)?)
            } else {
                Some(walk_statement_into_body(else_pair)?)
            }
        }
        None => None,
    };

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

/// Property names that are zero-arg getters in Dart but are bound as
/// value-method emitters in the profile (which only fires on Call).
/// Walker rewrites `expr.<name>` to `expr.<name>()` for these so the
/// dispatch path is uniform.
/// Build a Dart `expr is T` test. For primitive Dart types (int,
/// double, num, String, bool, List, Map, Object) we lower to a
/// `typeof`-style check via REF_TYPEOF so the test works on
/// primitives (which don't carry a `__type` field). Class types
/// fall back to the generic ExprKind::IsType which compares
/// `expr.__type == "T"`.
fn build_is_type(expr: Expression, type_name: &str) -> Expression {
    let raw = type_name.trim().trim_end_matches('?').trim();
    if raw == "List<int>" {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__dart_is_list_of_int")),
            args: vec![Argument::positional(expr)],
            optional: false,
        });
    }
    let trimmed = type_name
        .trim()
        .trim_end_matches('?')
        .split('<')
        .next()
        .unwrap_or(type_name.trim());
    if trimmed == "Null" {
        return Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(expr),
            right: Box::new(Expression::null()),
        });
    }
    if matches!(trimmed, "Map" | "Set" | "Record") {
        return Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(dart_runtime_type_name_expr(Expression::new(
                ExprKind::Call {
                    callee: Box::new(Expression::ident("__dart_runtime_type")),
                    args: vec![Argument::positional(expr)],
                    optional: false,
                },
            ))),
            right: Box::new(Expression::string(trimmed)),
        });
    }
    let typeof_tag: Option<&str> = match trimmed {
        "int" | "double" | "num" => Some("number"),
        "String" => Some("string"),
        "bool" => Some("boolean"),
        _ => None,
    };
    if let Some(tag) = typeof_tag {
        // Synthesise: `typeof expr === "<tag>"`.
        return Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(expr)))),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(tag.into())))),
        });
    }
    Expression::new(ExprKind::IsType {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

/// Dart property spellings this walker rewrites into zero-arg CALLS so the
/// `[value_methods]` dispatch table can see them.
///
/// The rewrite exists because a `[value_methods]` row is a CALL emitter, so a
/// property read has to be forged into a call to reach it — and for a BUILT-IN
/// receiver that is the only route there is. A built-in is CLASSIFIED, not
/// named: `var nums = [9, 9, 9]` infers to the array-shape spelling `int()`,
/// which matches no type node and never can, so every name-keyed path on the
/// member read (`lookup_type_property_target`, `lookup_type_instance_target`)
/// answers `None` for it. Its declared carrier is `[builtin_slots.<type>]`,
/// bound on the CALL path by `apply_builtin_slot_binding` (`calls.rs:5914`).
///
/// Forcing the call USED to be wrong once the receiver's class was known:
/// `user_typed_receiver_shadow` (`calls.rs:5993`) sent the call to the class,
/// the member get fired the getter and yielded the value, and the synthesised
/// call then called that value — `f64 is not callable (type: 9)`,
/// `bool is not callable (type: true)`.
///
/// **Declaring the tree leaf is what makes the force-call harmless**, and the
/// two are complementary carriers rather than duplicates. A `Property` leaf
/// serves BOTH paths: the member read consumes it at `expressions.rs:3081`,
/// and `lookup_type_instance_target` (`namespaces.rs:450`) unwraps the getter
/// for a zero-arg CALL — which resolves before the class shadow can divert it.
/// So a named receiver answers from the tree and a built-in answers from
/// `[value_methods]`, off the same forged call.
///
/// MEASURED, two rounds, failure sets stable: with `isEmpty`/`isNotEmpty` both
/// on this list and declared as `dart.core.StringBuffer` leaves, string_buffer
/// 41→48 and the collection slices hold at their baseline. Removing them from
/// this list instead cost `nums.isEmpty` and `m.isEmpty` on untyped literals.
///
/// `length` is OFF this list, and that is a known COMPROMISE. It survives only
/// because **`.length` is a native JS property on arrays and strings** — not
/// the tree, not a slot — so a receiver JS does not know silently answers
/// `undefined`. Putting it back (here + a `[value_methods] length` row, so it
/// reaches `emit_dart_length` and the `Len` slot) was MEASURED and reverted: it
/// breaks dart SETS, because a set literal is `__dart_set_from(Array([…]))`, a
/// CALL that cannot classify, and the flat `defined_class_methods` set then
/// claims `length` at `calls.rs:6017`. See the profile note for the full trace.
///
/// So no name comes off this list because it "reads like a property", and none
/// goes back on until its ROLE answers for EVERY receiver it can have.
fn is_dart_zero_arg_getter(name: &str) -> bool {
    matches!(
        name,
        "isEmpty"
            | "isNotEmpty"
            | "isEven"
            | "isOdd"
            | "isNegative"
            | "isNaN"
            | "isFinite"
            | "isInfinite"
            | "sign"
            | "first"
            | "last"
            | "single"
            | "singleOrNull"
            | "runes"
            | "codeUnits"
            | "keys"
            | "values"
            | "entries"
            | "reversed"
            | "isRunning"
            | "elapsed"
            | "elapsedMilliseconds"
            | "elapsedMicroseconds"
            // `Uri`'s property-shaped surface. Dart writes every one of these
            // without parentheses, and `core_classes/uri.rs` declares them as
            // zero-arg METHODS — because a property getter's body cannot see
            // `this` (measured on a plain user class, see that file). Without
            // the force-call a bare read yields the function object:
            // `[function authority]` instead of `example.com:8080`.
            //
            // All ten are Uri-specific in the dart suite, so force-calling them
            // on an arbitrary receiver has nothing else to hit.
            | "authority"
            | "origin"
            | "pathSegments"
            | "queryParameters"
            | "hasScheme"
            | "hasAuthority"
            | "hasQuery"
            | "hasFragment"
            | "hasEmptyPath"
            | "isAbsolute"
    )
}

// `dart_exception_constructor_alias` is GONE. It rewrote `Exception(msg)` into
// a `__dart_exception(msg)` builtin call, which built an anonymous struct with
// no rtt and a CANONICALIZED `__type` (`FormatException` → `ValueError`). The
// six types are `dart:core` CLASSES now (`core_classes/exceptions.rs`), so
// `dart_call_or_new` turns each construction into an ordinary `ExprKind::New`
// through `is_core_class` — the same path `StringBuffer(...)` takes.
/// Dart record positional field name `$1`/`$2`/… → its 0-based index. Records
/// are array-backed, so `rec.$1` lowers to `rec[0]`.
fn dart_positional_field_index(name: &str) -> Option<i64> {
    let n: i64 = name.strip_prefix('$')?.parse().ok()?;
    (n >= 1).then_some(n - 1)
}

/// Lower a list comprehension `[for (...) elem]` / `[if (...) elem]` to
/// an IIFE that builds the array imperatively. Walker-only normalization;
/// the compiler sees a regular Call(Lambda, []) on the way out.
fn lower_list_comprehension(elements: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let acc = "__compr_acc";
    let mut body: Vec<Statement> = Vec::new();
    // var __compr_acc = [];
    body.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(acc.to_string()),
            type_hint: None,
            init: Some(Expression::new(ExprKind::Array(Vec::new()))),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    }));
    for el in elements {
        body.push(lower_list_element(el, acc)?);
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Ident(acc.to_string()),
    )))));
    Ok(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

fn lower_list_element(el: Pair<Rule>, acc: &str) -> Result<Statement, String> {
    let inner = el.into_inner().next().ok_or("empty list element")?;
    match inner.as_rule() {
        Rule::collection_for => {
            // collection_for = "for" "(" for_header ")" list_element
            let mut header_pair = None;
            let mut body_pair = None;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::for_header => header_pair = Some(p),
                    Rule::list_element => body_pair = Some(p),
                    _ => {}
                }
            }
            let header = header_pair.ok_or("collection_for: missing header")?;
            let body_el = body_pair.ok_or("collection_for: missing body")?;
            let body_stmt = lower_list_element(body_el, acc)?;
            build_for_with_body(header, vec![body_stmt])
        }
        Rule::collection_if => {
            // collection_if = "if" "(" expression ")" list_element ("else" list_element)?
            let mut cond = None;
            let mut then_el = None;
            let mut else_el = None;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::expression if cond.is_none() => cond = Some(walk_expression(p)?),
                    Rule::list_element => {
                        if then_el.is_none() {
                            then_el = Some(p);
                        } else {
                            else_el = Some(p);
                        }
                    }
                    _ => {}
                }
            }
            let cond = cond.ok_or("collection_if: missing cond")?;
            let then_stmt = lower_list_element(then_el.ok_or("collection_if: missing then")?, acc)?;
            let else_stmt = match else_el {
                Some(el) => Some(vec![lower_list_element(el, acc)?]),
                None => None,
            };
            Ok(Statement::new(StmtKind::If {
                cond,
                then_body: vec![then_stmt],
                elifs: Vec::new(),
                else_body: else_stmt,
            }))
        }
        _ => {
            // Plain expression (or `... ~ expr` spread). Build `acc.add(expr)`.
            // Note: spread is not handled here yet — falls through as a single
            // value push (acceptable for compile_ok; runtime correctness for
            // spread inside comprehensions is a follow-up).
            let value = walk_expression(inner)?;
            let push_call = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::Ident(acc.to_string()))),
                    field: "add".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(value)],
                optional: false,
            });
            Ok(Statement::new(StmtKind::Expr(push_call)))
        }
    }
}

fn lower_set_comprehension(elements: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let acc = "__set_compr_acc";
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(acc.to_string()),
            type_hint: None,
            init: Some(Expression::new(ExprKind::Array(Vec::new()))),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    }));
    for el in elements {
        body.push(lower_set_element(el, acc)?);
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Ident(acc.to_string()),
    )))));
    Ok(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

fn lower_set_element(el: Pair<Rule>, acc: &str) -> Result<Statement, String> {
    let inner = el.into_inner().next().ok_or("empty set element")?;
    match inner.as_rule() {
        Rule::map_collection_for => {
            let mut header_pair = None;
            let mut body_pair = None;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::for_header => header_pair = Some(p),
                    Rule::map_or_set_element => body_pair = Some(p),
                    _ => {}
                }
            }
            let header = header_pair.ok_or("set collection_for: missing header")?;
            let body_el = body_pair.ok_or("set collection_for: missing body")?;
            let body_stmt = lower_set_element(body_el, acc)?;
            build_for_with_body(header, vec![body_stmt])
        }
        Rule::map_collection_if => {
            let mut cond = None;
            let mut then_el = None;
            let mut else_el = None;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::expression if cond.is_none() => cond = Some(walk_expression(p)?),
                    Rule::map_or_set_element => {
                        if then_el.is_none() {
                            then_el = Some(p);
                        } else {
                            else_el = Some(p);
                        }
                    }
                    _ => {}
                }
            }
            let cond = cond.ok_or("set collection_if: missing cond")?;
            let then_stmt =
                lower_set_element(then_el.ok_or("set collection_if: missing then")?, acc)?;
            let else_stmt = match else_el {
                Some(el) => Some(vec![lower_set_element(el, acc)?]),
                None => None,
            };
            Ok(Statement::new(StmtKind::If {
                cond,
                then_body: vec![then_stmt],
                elifs: Vec::new(),
                else_body: else_stmt,
            }))
        }
        _ => {
            let value = walk_expression(inner)?;
            let push_call = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::Ident(acc.to_string()))),
                    field: "add".to_string(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(value)],
                optional: false,
            });
            Ok(Statement::new(StmtKind::Expr(push_call)))
        }
    }
}

fn lower_map_comprehension(elements: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let acc = "__map_compr_acc";
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(acc.to_string()),
            type_hint: None,
            init: Some(Expression::new(ExprKind::Object(Vec::new()))),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    }));
    for el in elements {
        body.push(lower_map_element(el, acc)?);
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::ident(
        acc,
    )))));
    Ok(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

fn lower_map_element(el: Pair<Rule>, acc: &str) -> Result<Statement, String> {
    let inner = el.into_inner().next().ok_or("empty map element")?;
    match inner.as_rule() {
        Rule::map_collection_for => {
            let mut header_pair = None;
            let mut body_pair = None;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::for_header => header_pair = Some(p),
                    Rule::map_or_set_element => body_pair = Some(p),
                    _ => {}
                }
            }
            let header = header_pair.ok_or("map collection_for: missing header")?;
            let body_el = body_pair.ok_or("map collection_for: missing body")?;
            let body_stmt = lower_map_element(body_el, acc)?;
            build_for_with_body(header, vec![body_stmt])
        }
        Rule::map_collection_if => {
            let mut cond = None;
            let mut then_el = None;
            let mut else_el = None;
            for p in inner.into_inner() {
                match p.as_rule() {
                    Rule::expression if cond.is_none() => cond = Some(walk_expression(p)?),
                    Rule::map_or_set_element => {
                        if then_el.is_none() {
                            then_el = Some(p);
                        } else {
                            else_el = Some(p);
                        }
                    }
                    _ => {}
                }
            }
            let cond = cond.ok_or("map collection_if: missing cond")?;
            let then_stmt =
                lower_map_element(then_el.ok_or("map collection_if: missing then")?, acc)?;
            let else_stmt = match else_el {
                Some(el) => Some(vec![lower_map_element(el, acc)?]),
                None => None,
            };
            Ok(Statement::new(StmtKind::If {
                cond,
                then_body: vec![then_stmt],
                elifs: Vec::new(),
                else_body: else_stmt,
            }))
        }
        Rule::map_entry => {
            let mut parts = inner.into_inner();
            let key = walk_expression(parts.next().ok_or("map comprehension: missing key")?)?;
            let value = walk_expression(parts.next().ok_or("map comprehension: missing value")?)?;
            Ok(Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Assign {
                    target: Box::new(Expression::new(ExprKind::Index {
                        object: Box::new(Expression::ident(acc)),
                        index: Box::new(key),
                        null_safe: false,
                    })),
                    value: Box::new(value),
                },
            ))))
        }
        _ => Err(format!(
            "map comprehension: unexpected element {:?}",
            inner.as_rule()
        )),
    }
}

fn build_for_with_body(header_pair: Pair<Rule>, body: Vec<Statement>) -> Result<Statement, String> {
    let header_inner = header_pair.into_inner().next().ok_or("for: empty header")?;
    match header_inner.as_rule() {
        Rule::for_in_header => {
            let (var_name, iter, body) = lower_for_in_header_parts(header_inner, body)?;
            Ok(Statement::new(StmtKind::ForIn {
                var: var_name,
                key: None,
                iter,
                body,
                of: true,
                else_body: None,
                is_async: false,
            }))
        }
        Rule::for_c_header => {
            let mut init: Option<Box<Statement>> = None;
            let mut cond: Option<Expression> = None;
            let mut update: Option<Expression> = None;
            for p in header_inner.into_inner() {
                match p.as_rule() {
                    Rule::for_c_init => {
                        let inner = p.into_inner().next().ok_or("for init: empty")?;
                        match inner.as_rule() {
                            Rule::variable_declaration_no_semi => {
                                let stmt_kind = walk_var_decl_no_semi(inner)?;
                                init = Some(Box::new(Statement::new(stmt_kind)));
                            }
                            _ => {
                                let expr = walk_expression(inner)?;
                                init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                            }
                        }
                    }
                    Rule::expression => {
                        if cond.is_none() {
                            cond = Some(walk_expression(p)?);
                        }
                    }
                    Rule::for_c_update => {
                        let exprs: Result<Vec<Expression>, String> =
                            p.into_inner().map(walk_expression).collect();
                        let exprs = exprs?;
                        update = Some(if exprs.len() == 1 {
                            exprs.into_iter().next().unwrap()
                        } else {
                            Expression::new(ExprKind::Sequence(exprs))
                        });
                    }
                    _ => {}
                }
            }
            Ok(Statement::new(StmtKind::For {
                init,
                cond,
                update,
                body,
            }))
        }
        _ => Err(format!(
            "collection_for: unexpected header rule {:?}",
            header_inner.as_rule()
        )),
    }
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // for_statement = { "for" ~ "(" ~ for_header ~ ")" ~ statement }
    let mut header_pair = None;
    let mut body_pair = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::for_header => header_pair = Some(p),
            _ => body_pair = Some(p),
        }
    }

    let header = header_pair.ok_or("for: missing header")?;
    let body = walk_statement_into_body(body_pair.ok_or("for: missing body")?)?;

    let header_inner = header.into_inner().next().ok_or("for: empty header")?;

    match header_inner.as_rule() {
        Rule::for_in_header => {
            let (var_name, iter, body) = lower_for_in_header_parts(header_inner, body)?;
            Ok(StmtKind::ForIn {
                var: var_name,
                key: None,
                iter,
                body,
                of: true, // Dart for-in iterates values
                else_body: None,
                is_async: false,
            })
        }
        Rule::for_c_header => {
            // for_c_header = { for_c_init? ~ ";" ~ expression? ~ ";" ~ for_c_update? }
            let mut init: Option<Box<Statement>> = None;
            let mut cond: Option<Expression> = None;
            let mut update: Option<Expression> = None;

            for p in header_inner.into_inner() {
                match p.as_rule() {
                    Rule::for_c_init => {
                        let inner = p.into_inner().next().ok_or("for init: empty")?;
                        match inner.as_rule() {
                            Rule::variable_declaration_no_semi => {
                                let stmt_kind = walk_var_decl_no_semi(inner)?;
                                init = Some(Box::new(Statement::new(stmt_kind)));
                            }
                            _ => {
                                let expr = walk_expression(inner)?;
                                init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                            }
                        }
                    }
                    Rule::expression => {
                        if cond.is_none() {
                            cond = Some(walk_expression(p)?);
                        }
                    }
                    Rule::for_c_update => {
                        let exprs: Result<Vec<Expression>, String> =
                            p.into_inner().map(walk_expression).collect();
                        let exprs = exprs?;
                        update = Some(if exprs.len() == 1 {
                            exprs.into_iter().next().unwrap()
                        } else {
                            Expression::new(ExprKind::Sequence(exprs))
                        });
                    }
                    _ => {}
                }
            }

            Ok(StmtKind::For {
                init,
                cond,
                update,
                body,
            })
        }
        other => Err(format!("Unexpected for header: {:?}", other)),
    }
}

fn walk_var_decl_no_semi(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut var_kind = VarDeclKind::Let;
    let mut declarations = Vec::new();
    let mut type_hint: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_modifiers => {
                let txt = p.as_str().trim();
                if txt.contains("final") || txt.contains("const") {
                    var_kind = VarDeclKind::Const;
                }
            }
            Rule::type_or_var => {
                let inner_text = p.as_str().trim();
                if inner_text != "var" {
                    let has_var_kw = p.clone().into_inner().any(|c| c.as_rule() == Rule::var_kw);
                    if !has_var_kw {
                        type_hint = Some(inner_text.to_string());
                    }
                }
            }
            Rule::typed_var_declarator | Rule::var_declarator => {
                let decl = walk_var_declarator(p, type_hint.clone())?;
                declarations.push(decl);
            }
            _ => {}
        }
    }

    Ok(StmtKind::VarDecl {
        declarations,
        kind: var_kind,
    })
}

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(inner.next().ok_or("while: missing cond")?)?;
    let body = walk_statement_into_body(inner.next().ok_or("while: missing body")?)?;
    Ok(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
}

fn walk_do_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let body_pair = inner.next().ok_or("do-while: missing body")?;
    let body = walk_statement_into_body(body_pair)?;
    let cond = walk_expression(inner.next().ok_or("do-while: missing cond")?)?;
    Ok(StmtKind::DoWhile {
        body,
        cond,
        until: false,
    })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let subject = walk_expression(inner.next().ok_or("switch: missing expr")?)?;
    let subject_name = "__dart_switch_subject".to_string();
    let subject_expr = Expression::ident(&subject_name);
    let mut arms: Vec<(Expression, Vec<Statement>)> = Vec::new();
    let mut simple_cases: Vec<SwitchCase> = Vec::new();
    let mut default_body: Option<Vec<Statement>> = None;
    let mut needs_pattern_lowering = false;

    for p in inner {
        if p.as_rule() != Rule::switch_case {
            continue;
        }

        let src = p.as_str().trim_start();
        let is_default = src.starts_with("default");

        let mut children: Vec<Pair<Rule>> = p.into_inner().collect();

        if is_default {
            let stmts = children
                .into_iter()
                .filter_map(|c| walk_statement(c).transpose())
                .collect::<Result<Vec<_>, _>>()?;
            simple_cases.push(SwitchCase {
                conditions: vec![],
                body: stmts.clone(),
            });
            default_body = Some(stmts);
        } else {
            let pattern = children
                .iter()
                .position(|c| c.as_rule() == Rule::pattern)
                .map(|idx| children.remove(idx))
                .ok_or("switch: missing case pattern")?;
            let guard_idx = children
                .iter()
                .position(|c| c.as_rule() == Rule::when_guard);
            let simple_value = if guard_idx.is_none() {
                simple_switch_pattern_expr(pattern.clone())?
            } else {
                None
            };
            let original_pattern = pattern.clone();
            let mut analysis = analyze_dart_pattern(pattern, &subject_expr)?;
            if let Some(guard_idx) = guard_idx {
                let guard = children.remove(guard_idx);
                if let Some(guard_pair) = guard
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::conditional_expression)
                {
                    let guard_expr = substitute_pattern_bindings(
                        walk_expression(guard_pair)?,
                        &analysis.bindings,
                    );
                    analysis.cond = and_expr(analysis.cond, guard_expr);
                }
            }
            let mut stmts = children
                .into_iter()
                .filter_map(|c| walk_statement(c).transpose())
                .collect::<Result<Vec<_>, _>>()?;
            for stmt in &mut stmts {
                substitute_pattern_bindings_stmt(stmt, &analysis.bindings);
            }
            if let Some(value) = simple_value {
                simple_cases.push(SwitchCase {
                    conditions: vec![CaseCondition::Value(value.clone())],
                    body: stmts.clone(),
                });
            } else {
                needs_pattern_lowering = true;
                let _ = original_pattern;
            }
            arms.push((analysis.cond, stmts));
        }
    }

    if !needs_pattern_lowering {
        merge_empty_fallthrough_cases(&mut simple_cases);
        return Ok(StmtKind::Switch {
            expr: subject,
            cases: simple_cases,
            default: default_body,
        });
    }

    let mut else_body = default_body;
    for (cond, body) in arms.into_iter().rev() {
        let next = Statement::new(StmtKind::If {
            cond,
            then_body: body,
            elifs: Vec::new(),
            else_body,
        });
        else_body = Some(vec![next]);
    }

    let mut block = vec![Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(subject_name),
            type_hint: None,
            init: Some(subject),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })];
    if let Some(mut chain) = else_body {
        block.append(&mut chain);
    }
    Ok(StmtKind::Block(block))
}

fn simple_switch_pattern_expr(pair: Pair<Rule>) -> Result<Option<Expression>, String> {
    match pair.as_rule() {
        Rule::pattern => {
            let children: Vec<Pair<Rule>> = pair.into_inner().collect();
            if children.len() == 1 {
                simple_switch_pattern_expr(children.into_iter().next().unwrap())
            } else {
                Ok(None)
            }
        }
        Rule::primary_pattern => {
            let Some(inner) = pair.into_inner().next() else {
                return Ok(None);
            };
            simple_switch_pattern_expr(inner)
        }
        Rule::null_pattern => Ok(Some(Expression::null())),
        Rule::bool_pattern => Ok(Some(Expression::bool(pair.as_str().trim() == "true"))),
        Rule::signed_numeric_pattern => {
            let Some(n) = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::numeric_literal)
            else {
                return Ok(None);
            };
            let lit = Expression::new(walk_expr_kind(n)?);
            Ok(Some(Expression::new(ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(lit),
            })))
        }
        Rule::constant_pattern => {
            let children: Vec<Pair<Rule>> = pair.into_inner().collect();
            if children.len() != 1 {
                return Ok(None);
            }
            let child = children.into_iter().next().unwrap();
            match child.as_rule() {
                Rule::signed_numeric_pattern => simple_switch_pattern_expr(child),
                Rule::qualified_constant_pattern => {
                    let mut parts = child.into_inner();
                    let class = parts.next().ok_or("qualified pattern: missing class")?;
                    let member = parts.next().ok_or("qualified pattern: missing member")?;
                    Ok(Some(Expression::new(ExprKind::StaticAccess {
                        class: Box::new(Expression::ident(class.as_str())),
                        member: Box::new(Expression::ident(member.as_str())),
                    })))
                }
                Rule::numeric_literal | Rule::string_literal => {
                    Ok(Some(Expression::new(walk_expr_kind(child)?)))
                }
                Rule::ident_name => Ok(Some(Expression::ident(child.as_str()))),
                _ => Ok(None),
            }
        }
        _ => Ok(None),
    }
}

fn dart_stack_trace_binding(name: &str, caught: Option<&str>) -> Statement {
    let caught_expr = caught
        .map(Expression::ident)
        .unwrap_or_else(Expression::null);
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: Some("StackTrace".to_string().into()),
            init: Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__dart_stack_trace")),
                args: vec![Argument::positional(caught_expr)],
                optional: false,
            })),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

fn rewrite_direct_rethrow(body: &mut [Statement], caught: &str) {
    for stmt in body {
        if let StmtKind::Throw { expr, .. } = &mut stmt.kind {
            if expr.is_none() {
                *expr = Some(Expression::ident(caught));
            }
        }
    }
}

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally: Option<Vec<Statement>> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::block_statement => {
                if body.is_empty() {
                    body = walk_statement_into_body(p)?;
                }
            }
            Rule::catch_clause => {
                let inner = p.into_inner().next().ok_or("catch: empty")?;
                match inner.as_rule() {
                    Rule::on_catch_clause => {
                        // on Type catch (e, s) { }
                        let mut types = Vec::new();
                        let mut var_name: Option<String> = None;
                        let mut stack_var: Option<String> = None;
                        let mut catch_body = Vec::new();
                        let mut found_first_ident = false;

                        for cp in inner.into_inner() {
                            match cp.as_rule() {
                                Rule::ident_name => {
                                    if types.is_empty() && !found_first_ident {
                                        types.push(cp.as_str().to_string());
                                        found_first_ident = true;
                                    } else if var_name.is_none() {
                                        var_name = Some(cp.as_str().to_string());
                                    } else {
                                        stack_var = Some(cp.as_str().to_string());
                                    }
                                }
                                Rule::block_statement => {
                                    catch_body = walk_statement_into_body(cp)?;
                                }
                                _ => {}
                            }
                        }
                        if let Some(caught) = var_name.as_deref() {
                            rewrite_direct_rethrow(&mut catch_body, caught);
                        }
                        if let Some(stack_name) = stack_var.take() {
                            catch_body.insert(
                                0,
                                dart_stack_trace_binding(&stack_name, var_name.as_deref()),
                            );
                        }
                        catches.push(CatchClause {
                            types,
                            var_name,
                            stack_var,
                            body: catch_body,
                            when_clause: None,
                        });
                    }
                    Rule::plain_catch_clause => {
                        // catch (e, s) { }
                        let mut var_name: Option<String> = None;
                        let mut stack_var: Option<String> = None;
                        let mut catch_body = Vec::new();

                        for cp in inner.into_inner() {
                            match cp.as_rule() {
                                Rule::ident_name => {
                                    if var_name.is_none() {
                                        var_name = Some(cp.as_str().to_string());
                                    } else {
                                        stack_var = Some(cp.as_str().to_string());
                                    }
                                }
                                Rule::block_statement => {
                                    catch_body = walk_statement_into_body(cp)?;
                                }
                                _ => {}
                            }
                        }
                        if let Some(caught) = var_name.as_deref() {
                            rewrite_direct_rethrow(&mut catch_body, caught);
                        }
                        if let Some(stack_name) = stack_var.take() {
                            catch_body.insert(
                                0,
                                dart_stack_trace_binding(&stack_name, var_name.as_deref()),
                            );
                        }
                        catches.push(CatchClause {
                            types: Vec::new(),
                            var_name,
                            stack_var,
                            body: catch_body,
                            when_clause: None,
                        });
                    }
                    _ => {}
                }
            }
            Rule::finally_clause => {
                for fp in p.into_inner() {
                    if fp.as_rule() == Rule::block_statement {
                        finally = Some(walk_statement_into_body(fp)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    })
}

fn walk_yield_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let is_yield_from = pair.as_str().trim_start().starts_with("yield*");
    let mut value = None;
    for part in pair.into_inner() {
        if !is_kw(part.as_rule()) {
            value = Some(walk_expression(part)?);
        }
    }

    let expr = if is_yield_from {
        ExprKind::YieldFrom(Box::new(value.unwrap_or_else(Expression::null)))
    } else {
        ExprKind::Yield(value.map(Box::new))
    };
    Ok(StmtKind::Expr(Expression::new(expr)))
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(pair)?;
    Ok(normalize_dart_index_reads(
        Expression::with_span(kind, span),
        false,
    ))
}

fn dart_is_duration_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => matches!(name.as_str(), "Duration" | "Duration.zero"),
            ExprKind::Member { field, .. } => field == "elapsed",
            _ => false,
        },
        ExprKind::Member { object, field, .. } => {
            matches!(&object.kind, ExprKind::Ident(name) if name == "Duration") && field == "zero"
        }
        _ => false,
    }
}

fn dart_is_runtime_type_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__dart_runtime_type")
    )
}

fn dart_runtime_type_name_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__dart_type_to_string")),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
}

fn dart_type_literal_name(expr: &Expression) -> Option<String> {
    let ExprKind::Ident(name) = &expr.kind else {
        return None;
    };
    let base = name.split('<').next().unwrap_or(name).trim();
    if base.is_empty() || base.starts_with("__") {
        return None;
    }
    Some(base.to_string())
}

fn dart_runtime_type_compare_expr(left: Expression, right: Expression) -> Option<Expression> {
    if dart_is_runtime_type_expr(&left) {
        if let Some(type_name) = dart_type_literal_name(&right) {
            return Some(Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(dart_runtime_type_name_expr(left)),
                right: Box::new(Expression::string(&type_name)),
            }));
        }
    }
    if dart_is_runtime_type_expr(&right) {
        if let Some(type_name) = dart_type_literal_name(&left) {
            return Some(Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(dart_runtime_type_name_expr(right)),
                right: Box::new(Expression::string(&type_name)),
            }));
        }
    }
    None
}

fn dart_nullable_cast_expr(expr: Expression, type_name: &str) -> Expression {
    let base_type = type_name.trim().trim_end_matches('?').trim();
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(expr.clone()),
            right: Box::new(Expression::null()),
        })),
        then: Box::new(Expression::null()),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(build_is_type(expr.clone(), base_type)),
            then: Box::new(expr),
            else_: Box::new(Expression::null()),
        })),
    })
}

fn dart_duration_millis_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(expr),
        field: "inMilliseconds".to_string(),
        null_safe: false,
    })
}

fn normalize_dart_index_reads(expr: Expression, preserve_place: bool) -> Expression {
    let span = expr.span.clone();
    match expr.kind {
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => {
            let object = normalize_dart_index_reads(*object, false);
            let index = normalize_dart_index_reads(*index, false);
            if preserve_place {
                Expression::with_span(
                    ExprKind::Index {
                        object: Box::new(object),
                        index: Box::new(index),
                        null_safe,
                    },
                    span,
                )
            } else {
                Expression::with_span(
                    ExprKind::Call {
                        callee: Box::new(Expression::ident("__dart_index_get")),
                        args: vec![Argument::positional(object), Argument::positional(index)],
                        optional: false,
                    },
                    span,
                )
            }
        }
        ExprKind::Assign { target, value } => Expression::with_span(
            ExprKind::Assign {
                target: Box::new(normalize_dart_assignment_target(*target)),
                value: Box::new(normalize_dart_index_reads(*value, false)),
            },
            span,
        ),
        ExprKind::Binary { op, left, right } => Expression::with_span(
            match op {
                BinOp::Eq | BinOp::NotEq => {
                    let left = normalize_dart_index_reads(*left, false);
                    let right = normalize_dart_index_reads(*right, false);
                    let left_is_null = matches!(&left.kind, ExprKind::Lit(Literal::Null));
                    let right_is_null = matches!(&right.kind, ExprKind::Lit(Literal::Null));
                    let eq = if left_is_null || right_is_null {
                        // `x == null` / `null == x` is a reference null-test, not
                        // value equality — skip the deep `__dart_eq` cascade.
                        let operand = if left_is_null { right } else { left };
                        Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__dart_is_null")),
                            args: vec![Argument::positional(operand)],
                            optional: false,
                        })
                    } else if let Some(runtime_type_cmp) =
                        dart_runtime_type_compare_expr(left.clone(), right.clone())
                    {
                        runtime_type_cmp
                    } else if dart_is_duration_expr(&left) && dart_is_duration_expr(&right) {
                        Expression::new(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(dart_duration_millis_expr(left)),
                            right: Box::new(dart_duration_millis_expr(right)),
                        })
                    } else {
                        Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__dart_eq")),
                            args: vec![Argument::positional(left), Argument::positional(right)],
                            optional: false,
                        })
                    };
                    if op == BinOp::NotEq {
                        ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(eq),
                        }
                    } else {
                        eq.kind
                    }
                }
                BinOp::Add | BinOp::Sub => {
                    let left = normalize_dart_index_reads(*left, false);
                    let right = normalize_dart_index_reads(*right, false);
                    if dart_is_duration_expr(&left) && dart_is_duration_expr(&right) {
                        ExprKind::Call {
                            callee: Box::new(Expression::ident("Duration")),
                            args: vec![Argument::positional(Expression::new(ExprKind::Binary {
                                op,
                                left: Box::new(dart_duration_millis_expr(left)),
                                right: Box::new(dart_duration_millis_expr(right)),
                            }))],
                            optional: false,
                        }
                    } else {
                        ExprKind::Binary {
                            op,
                            left: Box::new(left),
                            right: Box::new(right),
                        }
                    }
                }
                _ => ExprKind::Binary {
                    op,
                    left: Box::new(normalize_dart_index_reads(*left, false)),
                    right: Box::new(normalize_dart_index_reads(*right, false)),
                },
            },
            span,
        ),
        ExprKind::Unary { op, expr } => Expression::with_span(
            ExprKind::Unary {
                op,
                expr: Box::new(normalize_dart_index_reads(*expr, false)),
            },
            span,
        ),
        ExprKind::Ternary { cond, then, else_ } => Expression::with_span(
            ExprKind::Ternary {
                cond: Box::new(normalize_dart_index_reads(*cond, false)),
                then: Box::new(normalize_dart_index_reads(*then, false)),
                else_: Box::new(normalize_dart_index_reads(*else_, false)),
            },
            span,
        ),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::with_span(
            ExprKind::Member {
                object: Box::new(normalize_dart_index_reads(*object, false)),
                field,
                null_safe,
            },
            span,
        ),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => Expression::with_span(
            ExprKind::Call {
                callee: Box::new(normalize_dart_index_reads(*callee, false)),
                args: args
                    .into_iter()
                    .map(|mut arg| {
                        arg.value = normalize_dart_index_reads(arg.value, false);
                        arg
                    })
                    .collect(),
                optional,
            },
            span,
        ),
        ExprKind::New { class, args } => Expression::with_span(
            ExprKind::New {
                class: Box::new(normalize_dart_index_reads(*class, false)),
                args: args
                    .into_iter()
                    .map(|mut arg| {
                        arg.value = normalize_dart_index_reads(arg.value, false);
                        arg
                    })
                    .collect(),
            },
            span,
        ),
        ExprKind::Array(items) => Expression::with_span(
            ExprKind::Array(
                items
                    .into_iter()
                    .map(|mut item| {
                        item.value = normalize_dart_index_reads(item.value, false);
                        item.key = item.key.map(|key| normalize_dart_index_reads(key, false));
                        item
                    })
                    .collect(),
            ),
            span,
        ),
        ExprKind::Object(items) => Expression::with_span(
            ExprKind::Object(
                items
                    .into_iter()
                    .map(|item| match item {
                        ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                            key: normalize_dart_index_reads(key, false),
                            value: normalize_dart_index_reads(value, false),
                        },
                        ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                            key: normalize_dart_index_reads(key, false),
                            value: normalize_dart_index_reads(value, false),
                        },
                        ObjectProperty::Spread(value) => {
                            ObjectProperty::Spread(normalize_dart_index_reads(value, false))
                        }
                        other => other,
                    })
                    .collect(),
            ),
            span,
        ),
        ExprKind::Tuple(items) => Expression::with_span(
            ExprKind::Tuple(
                items
                    .into_iter()
                    .map(|item| normalize_dart_index_reads(item, false))
                    .collect(),
            ),
            span,
        ),
        ExprKind::Await(inner) => Expression::with_span(
            ExprKind::Await(Box::new(normalize_dart_index_reads(*inner, false))),
            span,
        ),
        ExprKind::Yield(Some(inner)) => Expression::with_span(
            ExprKind::Yield(Some(Box::new(normalize_dart_index_reads(*inner, false)))),
            span,
        ),
        ExprKind::YieldFrom(inner) => Expression::with_span(
            ExprKind::YieldFrom(Box::new(normalize_dart_index_reads(*inner, false))),
            span,
        ),
        ExprKind::Cast { expr, type_name } => Expression::with_span(
            if type_name.trim().ends_with('?') {
                dart_nullable_cast_expr(normalize_dart_index_reads(*expr, false), &type_name).kind
            } else {
                ExprKind::Cast {
                    expr: Box::new(normalize_dart_index_reads(*expr, false)),
                    type_name,
                }
            },
            span,
        ),
        ExprKind::TypeOf(inner) => Expression::with_span(
            ExprKind::TypeOf(Box::new(normalize_dart_index_reads(*inner, false))),
            span,
        ),
        other => Expression::with_span(other, span),
    }
}

fn normalize_dart_assignment_target(expr: Expression) -> Expression {
    let span = expr.span.clone();
    match expr.kind {
        ExprKind::Call {
            callee,
            mut args,
            optional: false,
        } if matches!(&callee.kind, ExprKind::Ident(name) if name == "__dart_index_get")
            && args.len() == 2 =>
        {
            let index = args.pop().unwrap().value;
            let object = args.pop().unwrap().value;
            Expression::with_span(
                ExprKind::Index {
                    object: Box::new(normalize_dart_index_reads(object, false)),
                    index: Box::new(normalize_dart_index_reads(index, false)),
                    null_safe: false,
                },
                span,
            )
        }
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::with_span(
            ExprKind::Member {
                object: Box::new(normalize_dart_index_reads(*object, false)),
                field,
                null_safe,
            },
            span,
        ),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::with_span(
            ExprKind::Index {
                object: Box::new(normalize_dart_index_reads(*object, false)),
                index: Box::new(normalize_dart_index_reads(*index, false)),
                null_safe,
            },
            span,
        ),
        other => normalize_dart_index_reads(Expression::with_span(other, span), false),
    }
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        // ── Literals ────────────────────────────────────────────────────
        Rule::numeric_literal => {
            let s = pair.as_str().replace('_', "");
            // Radix prefixes MUST be tested before the float test: a hex digit
            // can legitimately be `e`/`E` (`0xe000`, `0xFE`), and the exponent
            // check would otherwise claim those as malformed floats.
            if let Some(hex) = s.strip_prefix("0x").or_else(|| s.strip_prefix("0X")) {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(hex, 16).map_err(|e| format!("{}", e))?,
                )))
            } else if let Some(bin) = s.strip_prefix("0b").or_else(|| s.strip_prefix("0B")) {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(bin, 2).map_err(|e| format!("{}", e))?,
                )))
            } else if s.contains('.') || s.contains('e') || s.contains('E') {
                Ok(ExprKind::Lit(Literal::Float(
                    s.parse().map_err(|e| format!("{}", e))?,
                )))
            } else {
                Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
            }
        }

        // ── String literals ─────────────────────────────────────────────
        Rule::string_literal_sequence => {
            let mut parts = Vec::new();
            let mut has_interp = false;
            for inner in pair.into_inner() {
                let literal = if inner.as_rule() == Rule::string_literal {
                    inner.into_inner().next().ok_or("empty string")?
                } else {
                    inner
                };
                let expr = walk_string_literal(literal)?;
                match expr {
                    ExprKind::Lit(Literal::Str(text)) => parts.push(InterpolPart::Text(text)),
                    ExprKind::Interpolation(mut nested) => {
                        has_interp = true;
                        parts.append(&mut nested);
                    }
                    other => parts.push(InterpolPart::Expr(Expression::new(other))),
                }
            }
            if has_interp {
                Ok(ExprKind::Interpolation(parts))
            } else {
                Ok(ExprKind::Lit(Literal::Str(
                    parts
                        .into_iter()
                        .filter_map(|part| match part {
                            InterpolPart::Text(text) => Some(text),
                            _ => None,
                        })
                        .collect(),
                )))
            }
        }
        Rule::string_literal => {
            let inner = pair.into_inner().next().ok_or("empty string")?;
            walk_string_literal(inner)
        }

        // `#name` is a Symbol, and Dart renders a Symbol as `Symbol("name")` —
        // that rendering IS the value here, so `#mode` compares equal to an
        // `Invocation.memberName` and indexes `namedArguments` without a
        // separate Symbol value type. Carrying the raw `#name` text instead
        // made both of those silently miss.
        Rule::symbol_literal => Ok(ExprKind::Lit(Literal::Str(nsm_symbol(
            pair.as_str().trim_start_matches('#'),
        )))),

        Rule::raw_string => {
            let s = pair.as_str();
            // r'...' or r"..."
            let inner = if s.starts_with("r'") {
                &s[2..s.len() - 1]
            } else {
                &s[2..s.len() - 1]
            };
            Ok(ExprKind::Lit(Literal::Str(inner.to_string())))
        }

        // Interpolated strings
        Rule::interpolated_double_string | Rule::interpolated_single_string => {
            walk_interpolated_string(pair)
        }

        Rule::triple_double_string | Rule::triple_single_string => walk_interpolated_string(pair),

        // ── Keywords ────────────────────────────────────────────────────
        Rule::this_kw => Ok(ExprKind::This),
        Rule::super_kw => Ok(ExprKind::Super),
        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::null_kw => Ok(ExprKind::Lit(Literal::Null)),

        // ── Identifiers ─────────────────────────────────────────────────
        Rule::ident_name => {
            let name = pair.as_str();
            Ok(ExprKind::Ident(name.to_string()))
        }

        Rule::typed_ident => {
            // `Stream<int>` — keep just the identifier; type args are erased.
            let mut inner = pair.into_inner();
            let name_pair = inner.next().ok_or("typed_ident: missing name")?;
            Ok(ExprKind::Ident(name_pair.as_str().to_string()))
        }

        // ── Comma expression ────────────────────────────────────────────
        Rule::expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else {
                let exprs: Vec<Expression> = inner
                    .into_iter()
                    .map(walk_expression)
                    .collect::<Result<Vec<_>, _>>()?;
                if exprs.len() == 1 {
                    Ok(exprs.into_iter().next().unwrap().kind)
                } else {
                    Ok(ExprKind::Sequence(exprs))
                }
            }
        }

        // ── Assignment expression ───────────────────────────────────────
        Rule::assignment_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else if inner.len() >= 2 {
                // Check if first child is lambda_expression
                if inner[0].as_rule() == Rule::lambda_expression {
                    return walk_expr_kind(inner.remove(0));
                }
                // conditional_expression ~ (assignment_op ~ assignment_expression)?
                let left_pair = inner.remove(0);
                let left_span = to_span(&left_pair);
                let left = Expression::with_span(walk_expr_kind(left_pair)?, left_span);
                if inner.is_empty() {
                    return Ok(left.kind);
                }
                let op_str = inner.remove(0).as_str().trim();
                let right = walk_expression(inner.remove(0))?;

                if op_str == "=" {
                    Ok(ExprKind::Assign {
                        target: Box::new(left),
                        value: Box::new(right),
                    })
                } else {
                    let op = match op_str {
                        "+=" => CompoundOp::Add,
                        "-=" => CompoundOp::Sub,
                        "*=" => CompoundOp::Mul,
                        "/=" => CompoundOp::Div,
                        "~/=" => CompoundOp::IDiv,
                        "%=" => CompoundOp::Mod,
                        "&=" => CompoundOp::BitAnd,
                        "|=" => CompoundOp::BitOr,
                        "^=" => CompoundOp::BitXor,
                        "<<=" => CompoundOp::Shl,
                        ">>=" => CompoundOp::Shr,
                        ">>>=" => CompoundOp::UShr,
                        "??=" => CompoundOp::NullCoalesce,
                        _ => CompoundOp::Add,
                    };
                    Ok(ExprKind::Assign {
                        target: Box::new(left.clone()),
                        value: Box::new(Expression::new(ExprKind::Binary {
                            op: compound_to_binop(op),
                            left: Box::new(left),
                            right: Box::new(right),
                        })),
                    })
                }
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }

        Rule::throw_expression => Ok(walk_throw_expression(pair)?.kind),

        // ── Lambda / arrow function ─────────────────────────────────────
        Rule::lambda_expression => {
            let mut params = Vec::new();
            let mut body = LambdaBody::Expr(Box::new(Expression::null()));
            let mut is_async = false;

            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::async_kw => is_async = true,
                    Rule::lambda_params => {
                        for lp in p.into_inner() {
                            match lp.as_rule() {
                                Rule::lambda_param_list => {
                                    for item in lp.into_inner() {
                                        match item.as_rule() {
                                            Rule::lambda_param => {
                                                params.push(walk_lambda_param_pair(item)?);
                                            }
                                            Rule::lambda_param_item => {
                                                for inner in item.into_inner() {
                                                    match inner.as_rule() {
                                                        Rule::lambda_param => {
                                                            params.push(walk_lambda_param_pair(
                                                                inner,
                                                            )?);
                                                        }
                                                        Rule::named_lambda_params => {
                                                            for named in inner.into_inner() {
                                                                if named.as_rule()
                                                                    == Rule::lambda_param
                                                                {
                                                                    let mut param =
                                                                        walk_lambda_param_pair(
                                                                            named,
                                                                        )?;
                                                                    param.is_optional = true;
                                                                    params.push(param);
                                                                }
                                                            }
                                                        }
                                                        _ => {}
                                                    }
                                                }
                                            }
                                            Rule::named_lambda_params => {
                                                for named in item.into_inner() {
                                                    if named.as_rule() == Rule::lambda_param {
                                                        let mut param =
                                                            walk_lambda_param_pair(named)?;
                                                        param.is_optional = true;
                                                        params.push(param);
                                                    }
                                                }
                                            }
                                            _ => {}
                                        }
                                    }
                                }
                                Rule::ident_name => {
                                    params = vec![Param {
                                        name: lp.as_str().to_string(),
                                        type_hint: None,
                                        default: None,
                                        pass_by: PassBy::Value,
                                        is_rest: false,
                                        is_kwargs: false,
                                        is_optional: false,
                                        is_nullable: false,
                                    }];
                                }
                                _ => {}
                            }
                        }
                    }
                    Rule::arrow_op => {}
                    Rule::throw_expression => {
                        body = LambdaBody::Block(vec![Statement::new(StmtKind::Throw {
                            expr: Some(walk_throw_expression(p)?),
                            cause: None,
                        })]);
                    }
                    Rule::assignment_expression => {
                        body = LambdaBody::Expr(Box::new(walk_expression(p)?));
                    }
                    Rule::function_body_block => {
                        body = LambdaBody::Block(walk_statement_into_body(p)?);
                    }
                    _ => {}
                }
            }

            Ok(ExprKind::Lambda {
                params,
                body,
                is_async,
                captures: Vec::new(),
            })
        }

        // ── Ternary / conditional ───────────────────────────────────────
        Rule::conditional_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else if inner.len() >= 3 {
                // null_coalesce ~ "?" ~ expression ~ ":" ~ expression
                let cond = walk_expression(inner.remove(0))?;
                let then = walk_expression(inner.remove(0))?;
                let else_ = walk_expression(inner.remove(0))?;
                Ok(ExprKind::Ternary {
                    cond: Box::new(cond),
                    then: Box::new(then),
                    else_: Box::new(else_),
                })
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }

        // ── Binary expression (flat Pratt) ──────────────────────────────
        //
        // Grammar collapses 12 precedence layers into one `(operand ~
        // (op ~ operand)*)` rule. The walker climbs the resulting flat
        // sequence into a precedence-correct tree using the standard
        // shunting-yard algorithm.
        Rule::null_coalesce_expression => {
            let inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            walk_pratt(inner)
        }

        // ── Relational unit (unary + is/as/relational suffixes) ─────────
        Rule::relational_unit => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                return walk_expr_kind(inner.remove(0));
            }
            let mut left = walk_expression(inner.remove(0))?;
            for p in inner {
                if p.as_rule() != Rule::relational_suffix {
                    continue;
                }
                let mut children: Vec<Pair<Rule>> = p.into_inner().collect();
                let first = children.remove(0);
                match first.as_rule() {
                    Rule::is_test => {
                        let type_name = extract_type_from_inner(first);
                        left = build_is_type(left, &type_name);
                    }
                    Rule::is_not_test => {
                        let type_name = extract_type_from_inner(first);
                        left = Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(build_is_type(left, &type_name)),
                        });
                    }
                    Rule::as_cast => {
                        let type_name = extract_type_from_inner(first);
                        left = Expression::new(ExprKind::Cast {
                            expr: Box::new(left),
                            type_name,
                        });
                    }
                    Rule::relational_op => {
                        let op_str = first.as_str().trim();
                        let right = walk_expression(children.remove(0))?;
                        let op = match op_str {
                            "<" => BinOp::Lt,
                            ">" => BinOp::Gt,
                            "<=" => BinOp::LtEq,
                            ">=" => BinOp::GtEq,
                            _ => BinOp::Lt,
                        };
                        left = Expression::new(ExprKind::Binary {
                            op,
                            left: Box::new(left),
                            right: Box::new(right),
                        });
                    }
                    _ => {}
                }
            }
            Ok(left.kind)
        }

        // ── Unary ───────────────────────────────────────────────────────
        Rule::unary_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                return walk_expr_kind(inner.remove(0));
            }
            // unary_op ~ unary_expression
            let first = inner.remove(0);
            if first.as_rule() == Rule::unary_op {
                let op_str = first.as_str().trim();
                let operand = walk_expression(inner.remove(0))?;
                if op_str.starts_with("await") {
                    return Ok(ExprKind::Await(Box::new(operand)));
                }
                let op = match op_str {
                    "-" => UnaryOp::Neg,
                    "!" => UnaryOp::Not,
                    "~" => UnaryOp::BitNot,
                    "++" => UnaryOp::PreInc,
                    "--" => UnaryOp::PreDec,
                    _ => UnaryOp::Neg,
                };
                Ok(ExprKind::Unary {
                    op,
                    expr: Box::new(operand),
                })
            } else {
                // postfix_expression fallthrough
                walk_expr_kind(first)
            }
        }

        Rule::unary_op => {
            // Should not be reached directly
            Err(format!("unexpected bare unary_op: {}", pair.as_str()))
        }

        // ── Postfix ─────────────────────────────────────────────────────
        Rule::postfix_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            let base = walk_expression(inner.remove(0))?;
            if let Some(postfix) = inner.iter().find(|p| p.as_rule() == Rule::postfix_op) {
                let op = match postfix.as_str() {
                    "++" => UnaryOp::PostInc,
                    "--" => UnaryOp::PostDec,
                    _ => return Ok(base.kind),
                };
                Ok(ExprKind::Unary {
                    op,
                    expr: Box::new(base),
                })
            } else {
                Ok(base.kind)
            }
        }

        // ── Call / member / index chain ─────────────────────────────────
        Rule::call_expression => walk_call_chain(pair),

        Rule::new_expression => {
            let inner = pair.into_inner();
            // new_expression = { "new" ~ ident_name ~ ("." ~ ident_name)? ~ type_args? ~ "(" ~ argument_list? ~ ")" }
            let mut class_parts: Vec<String> = Vec::new();
            let mut args = Vec::new();
            for p in inner {
                match p.as_rule() {
                    Rule::ident_name => class_parts.push(p.as_str().to_string()),
                    Rule::type_args => {}
                    Rule::argument_list => args = walk_arguments(p)?,
                    _ => {}
                }
            }
            let class_name = class_parts.join(".");
            if let Some((ty, ctor)) = class_name.split_once('.') {
                if let Some(kind) = dart_flutter_named_ctor(ty, ctor, &args) {
                    return Ok(kind);
                }
            }
            inject_flutter_defaults(&class_name, &mut args);
            Ok(ExprKind::New {
                class: Box::new(Expression::ident(&class_name)),
                args,
            })
        }

        Rule::const_expression => {
            // Dart CANONICALIZES const values: two const expressions that
            // denote the same value ARE the same object, which is what makes
            // `identical(const [1, 2], const [1, 2])` true. Lowering each
            // occurrence inline built a fresh object every time, so every
            // `identical` over consts answered false.
            //
            // So each distinct const expression is built ONCE, hoisted to a
            // top-level binding, and every occurrence becomes a reference to
            // it. Identity of const expressions is keyed on their source text
            // with whitespace collapsed — `const [1,2]` and `const [1, 2]`
            // canonicalize together. Two spellings of the same VALUE
            // (`const Token(1)` vs `const Token(0 + 1)`) do not; Dart would
            // canonicalize those too, and that needs const evaluation.
            if let Some(name) = dart_const_pool_lookup(pair.as_str()) {
                return Ok(ExprKind::Ident(name));
            }
            let key = pair.as_str().to_string();
            let lowered = Expression::new(walk_const_expression_value(pair)?);
            return Ok(ExprKind::Ident(dart_const_pool_insert(key, lowered)));
        }

        // ── Primary ─────────────────────────────────────────────────────
        Rule::primary => {
            let inner = pair.into_inner().next().ok_or("empty primary")?;
            walk_expr_kind(inner)
        }

        // ── Switch expression (Dart 3) ──────────────────────────────────
        Rule::switch_expression => {
            let mut inner = pair.into_inner();
            let subject = walk_expression(inner.next().ok_or("switch expr: missing subject")?)?;
            let mut arms: Vec<(Expression, Expression)> = Vec::new();
            for p in inner {
                if p.as_rule() == Rule::switch_expr_case {
                    let mut case_inner = p.into_inner();
                    let pattern = case_inner.next().ok_or("switch expr: missing pattern")?;
                    let mut analysis = analyze_dart_pattern(pattern, &subject)?;
                    let mut body_expr = None;
                    for cp in case_inner {
                        match cp.as_rule() {
                            Rule::when_guard => {
                                if let Some(guard_pair) = cp
                                    .into_inner()
                                    .find(|p| p.as_rule() == Rule::conditional_expression)
                                {
                                    let guard = substitute_pattern_bindings(
                                        walk_expression(guard_pair)?,
                                        &analysis.bindings,
                                    );
                                    analysis.cond = and_expr(analysis.cond, guard);
                                }
                            }
                            Rule::assignment_expression => {
                                body_expr = Some(substitute_pattern_bindings(
                                    walk_expression(cp)?,
                                    &analysis.bindings,
                                ));
                            }
                            _ => {}
                        }
                    }
                    if let Some(body) = body_expr {
                        arms.push((analysis.cond, body));
                    }
                }
            }
            Ok(lower_switch_expr_arms(arms).kind)
        }

        // ── Paren / record expression ───────────────────────────────────
        Rule::record_or_paren => {
            // A lone trailing comma is what makes `(99,)` a one-field record
            // rather than the grouping `(99)`; the comma is dropped by the
            // grammar, so read it off the source.
            let single_field_record = pair
                .as_str()
                .trim_end()
                .trim_end_matches(')')
                .trim_end()
                .ends_with(',');
            let inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.is_empty() {
                // () — empty tuple/record
                return Ok(ExprKind::Tuple(Vec::new()));
            }
            // record_or_paren_inner contains record_fields
            let ropi = inner.into_iter().next().unwrap();
            if ropi.as_rule() == Rule::record_or_paren_inner {
                let fields: Vec<Pair<Rule>> = ropi
                    .into_inner()
                    .filter(|p| p.as_rule() == Rule::record_field)
                    .collect();

                // Check if any field has a name label (record) or single expression (paren)
                if fields.len() == 1 {
                    let field_children: Vec<Pair<Rule>> =
                        fields.into_iter().next().unwrap().into_inner().collect();
                    if field_children.len() == 1 {
                        let value = walk_expression(field_children.into_iter().next().unwrap())?;
                        // `(x,)` is a one-element record; `(x)` is just grouping.
                        return Ok(if single_field_record {
                            ExprKind::Tuple(vec![value])
                        } else {
                            value.kind
                        });
                    } else if field_children.len() == 2 {
                        // Single named field record `(name: value)` — the shared
                        // canonical named-tuple shape (array-backed + by-name key).
                        let name = field_children[0].as_str().to_string();
                        let value = walk_expression(field_children.into_iter().nth(1).unwrap())?;
                        return Ok(vybe_compiler::primitives::tuples::build_named_tuple(vec![
                            (Some(name), value),
                        ]));
                    }
                    // Fallthrough — treat as empty
                    return Ok(ExprKind::Lit(Literal::Null));
                }

                // Multiple fields — could be record or tuple
                let has_named = fields.iter().any(|f| f.clone().into_inner().count() > 1);
                if has_named {
                    // Mixed/named record → the shared canonical named-tuple shape:
                    // array-backed (so `.$1`/`.$2` index positionally) with a
                    // by-name key per labelled field (`.host`, `.port`). One value
                    // across languages (Python namedtuple / C# ValueTuple).
                    let mut record_fields: Vec<(Option<String>, Expression)> = Vec::new();
                    for f in fields {
                        let mut fi = f.into_inner();
                        let first = fi.next().unwrap();
                        if let Some(second) = fi.next() {
                            let key = first.as_str().to_string();
                            record_fields.push((Some(key), walk_expression(second)?));
                        } else {
                            record_fields.push((None, walk_expression(first)?));
                        }
                    }
                    Ok(vybe_compiler::primitives::tuples::build_named_tuple(
                        record_fields,
                    ))
                } else {
                    let exprs: Vec<Expression> = fields
                        .into_iter()
                        .map(|f| walk_expression(f.into_inner().next().unwrap()))
                        .collect::<Result<Vec<_>, _>>()?;
                    if exprs.len() == 1 {
                        Ok(exprs.into_iter().next().unwrap().kind)
                    } else {
                        Ok(ExprKind::Tuple(exprs))
                    }
                }
            } else {
                walk_expr_kind(ropi)
            }
        }

        // ── List literal ────────────────────────────────────────────────
        Rule::list_literal => {
            // If any element is a collection-for or collection-if, lower the
            // whole list to an IIFE that builds it imperatively. Otherwise
            // emit a plain array literal.
            let elements: Vec<Pair<Rule>> = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::list_element)
                .collect();
            let has_comprehension = elements.iter().any(|p| {
                p.clone()
                    .into_inner()
                    .next()
                    .map(|c| matches!(c.as_rule(), Rule::collection_for | Rule::collection_if))
                    .unwrap_or(false)
            });
            if has_comprehension {
                return Ok(lower_list_comprehension(elements)?);
            }
            let mut out = Vec::new();
            for p in elements {
                let src = p.as_str().trim_start();
                let spread = src.starts_with("...");
                let inner = p
                    .into_inner()
                    .next()
                    .ok_or("empty list element".to_string())?;
                let value = walk_expression(inner)?;
                out.push(ArrayElement {
                    key: None,
                    value,
                    spread,
                    by_ref: false,
                });
            }
            Ok(ExprKind::Array(out))
        }

        // ── Map / set literal ───────────────────────────────────────────
        Rule::map_or_set_literal => {
            let mut props = Vec::new();
            let mut is_set = false;
            let mut is_map = false;
            let mut elements = Vec::new();

            fn walk_one(
                elem: Pair<Rule>,
                props: &mut Vec<ObjectProperty>,
                is_map: &mut bool,
                is_set: &mut bool,
            ) -> Result<(), String> {
                match elem.as_rule() {
                    Rule::map_or_set_element => {
                        let src = elem.as_str().trim_start();
                        if src.starts_with("...") {
                            if let Some(value_pair) = elem
                                .clone()
                                .into_inner()
                                .find(|p| p.as_rule() == Rule::assignment_expression)
                            {
                                let value = walk_expression(value_pair)?;
                                props.push(ObjectProperty::Spread(value));
                                return Ok(());
                            }
                        }
                        for inner in elem.into_inner() {
                            walk_one(inner, props, is_map, is_set)?;
                        }
                    }
                    Rule::map_entry => {
                        *is_map = true;
                        let mut ei = elem.into_inner();
                        let key = walk_expression(ei.next().ok_or("map entry: no key")?)?;
                        let value = walk_expression(ei.next().ok_or("map entry: no value")?)?;
                        props.push(ObjectProperty::KeyValue { key, value });
                    }
                    Rule::assignment_expression => {
                        *is_set = true;
                        let value = walk_expression(elem)?;
                        props.push(ObjectProperty::KeyValue {
                            key: Expression::null(),
                            value,
                        });
                    }
                    Rule::map_collection_if | Rule::map_collection_for => {
                        if elem.as_str().contains(':') {
                            *is_map = true;
                        } else {
                            *is_set = true;
                        }
                    }
                    // spread — skip body for now; grammar tolerates it so source compiles.
                    _ => {}
                }
                Ok(())
            }

            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::type_args => {
                        if p.as_str().contains(',') {
                            is_map = true;
                        } else {
                            is_set = true;
                        }
                    }
                    Rule::map_or_set_body => {
                        for entry in p.into_inner() {
                            elements.push(entry.clone());
                            walk_one(entry, &mut props, &mut is_map, &mut is_set)?;
                        }
                    }
                    _ => {}
                }
            }

            let has_comprehension = elements.iter().any(|p| {
                p.clone()
                    .into_inner()
                    .next()
                    .map(|c| {
                        matches!(
                            c.as_rule(),
                            Rule::map_collection_for | Rule::map_collection_if
                        )
                    })
                    .unwrap_or(false)
            });
            if is_map && has_comprehension {
                return Ok(lower_map_comprehension(elements)?);
            }
            if is_set && !is_map {
                let has_comprehension = elements.iter().any(|p| {
                    p.clone()
                        .into_inner()
                        .next()
                        .map(|c| {
                            matches!(
                                c.as_rule(),
                                Rule::map_collection_for | Rule::map_collection_if
                            )
                        })
                        .unwrap_or(false)
                });
                if has_comprehension {
                    return Ok(ExprKind::Call {
                        callee: Box::new(Expression::ident("__dart_set_from")),
                        args: vec![Argument::positional(Expression::new(
                            lower_set_comprehension(elements)?,
                        ))],
                        optional: false,
                    });
                }
                let elements: Vec<ArrayElement> = props
                    .into_iter()
                    .filter_map(|p| match p {
                        ObjectProperty::KeyValue { value, .. } => Some(ArrayElement {
                            key: None,
                            value,
                            spread: false,
                            by_ref: false,
                        }),
                        ObjectProperty::Spread(value) => Some(ArrayElement {
                            key: None,
                            value,
                            spread: true,
                            by_ref: false,
                        }),
                        _ => None,
                    })
                    .collect();
                Ok(ExprKind::Call {
                    callee: Box::new(Expression::ident("__dart_set_from")),
                    args: vec![Argument::positional(Expression::new(ExprKind::Array(
                        elements,
                    )))],
                    optional: false,
                })
            } else {
                Ok(ExprKind::Object(props))
            }
        }

        // ── Passthrough wrappers ────────────────────────────────────────
        Rule::call_chain => {
            let inner = pair.into_inner().next().ok_or("empty call_chain")?;
            walk_expr_kind(inner)
        }

        other => Err(format!(
            "Dart walker: unexpected expression rule: {:?} = {:?}",
            other,
            pair.as_str()
        )),
    }
}

struct DartPatternAnalysis {
    cond: Expression,
    bindings: HashMap<String, Expression>,
}

fn lower_switch_expr_arms(arms: Vec<(Expression, Expression)>) -> Expression {
    let mut fallback = Expression::null();
    for (cond, body) in arms.into_iter().rev() {
        fallback = Expression::new(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(body),
            else_: Box::new(fallback),
        });
    }
    fallback
}

fn analyze_dart_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    match pair.as_rule() {
        Rule::pattern => {
            let mut inner = pair.into_inner();
            let first = inner.next().ok_or("pattern: empty")?;
            let mut acc = analyze_dart_pattern(first, subject)?;
            while let Some(op) = inner.next() {
                let rhs_pair = inner.next().ok_or("pattern: missing rhs")?;
                let rhs = analyze_dart_pattern(rhs_pair, subject)?;
                acc.cond = match op.as_str() {
                    "&&" => and_expr(acc.cond, rhs.cond),
                    _ => or_expr(acc.cond, rhs.cond),
                };
                acc.bindings.extend(rhs.bindings);
            }
            Ok(acc)
        }
        Rule::primary_pattern => {
            let inner = pair.into_inner().next().ok_or("primary pattern: empty")?;
            analyze_dart_pattern(inner, subject)
        }
        Rule::wildcard_pattern => Ok(pattern_cond(Expression::bool(true))),
        Rule::variable_pattern => {
            let mut bindings = HashMap::new();
            if let Some(name) = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
            {
                if name != "_" {
                    bindings.insert(name, subject.clone());
                }
            }
            Ok(DartPatternAnalysis {
                cond: Expression::bool(true),
                bindings,
            })
        }
        Rule::null_pattern => Ok(pattern_cond(eq_expr(subject.clone(), Expression::null()))),
        Rule::bool_pattern => {
            let value = pair.as_str().trim() == "true";
            Ok(pattern_cond(eq_expr(
                subject.clone(),
                Expression::bool(value),
            )))
        }
        Rule::constant_pattern => analyze_constant_pattern(pair, subject),
        Rule::signed_numeric_pattern => {
            let n = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::numeric_literal)
                .ok_or("signed numeric pattern: missing literal")?;
            let lit = Expression::new(walk_expr_kind(n)?);
            Ok(pattern_cond(eq_expr(
                subject.clone(),
                Expression::new(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(lit),
                }),
            )))
        }
        Rule::relational_pattern => {
            let op_src = pair.as_str().trim_start();
            let op = if op_src.starts_with("<=") {
                BinOp::LtEq
            } else if op_src.starts_with(">=") {
                BinOp::GtEq
            } else if op_src.starts_with("==") {
                BinOp::Eq
            } else if op_src.starts_with("!=") {
                BinOp::NotEq
            } else if op_src.starts_with('<') {
                BinOp::Lt
            } else {
                BinOp::Gt
            };
            let rhs = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::assignment_expression)
                .map(walk_expression)
                .transpose()?
                .unwrap_or_else(Expression::null);
            Ok(pattern_cond(Expression::new(ExprKind::Binary {
                op,
                left: Box::new(subject.clone()),
                right: Box::new(rhs),
            })))
        }
        Rule::list_pattern => analyze_list_pattern(pair, subject),
        Rule::map_pattern => analyze_map_pattern(pair, subject),
        Rule::record_pattern => analyze_record_pattern(pair, subject),
        Rule::object_pattern => analyze_object_pattern(pair, subject),
        _ => Ok(pattern_cond(eq_expr(
            subject.clone(),
            walk_expression(pair)?,
        ))),
    }
}

fn analyze_constant_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    if children.len() == 2
        && children.iter().all(|p| p.as_rule() == Rule::ident_name)
        && children[0]
            .as_str()
            .chars()
            .next()
            .map(|ch| ch.is_ascii_uppercase())
            .unwrap_or(false)
    {
        let mut bindings = HashMap::new();
        let name = children[1].as_str().to_string();
        if name != "_" {
            bindings.insert(name, subject.clone());
        }
        return Ok(DartPatternAnalysis {
            cond: Expression::bool(true),
            bindings,
        });
    }
    let value = children
        .into_iter()
        .next()
        .map(|child| {
            if child.as_rule() == Rule::signed_numeric_pattern {
                let n = child
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::numeric_literal)
                    .ok_or("signed numeric pattern: missing literal")?;
                let lit = Expression::new(walk_expr_kind(n)?);
                Ok(Expression::new(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(lit),
                }))
            } else if child.as_rule() == Rule::qualified_constant_pattern {
                let mut parts = child.into_inner();
                let class = parts.next().ok_or("qualified pattern: missing class")?;
                let member = parts.next().ok_or("qualified pattern: missing member")?;
                Ok(Expression::new(ExprKind::StaticAccess {
                    class: Box::new(Expression::ident(class.as_str())),
                    member: Box::new(Expression::ident(member.as_str())),
                }))
            } else {
                walk_expression(child)
            }
        })
        .transpose()?
        .unwrap_or_else(Expression::null);
    Ok(pattern_cond(eq_expr(subject.clone(), value)))
}

fn analyze_list_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    let elements: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::list_pattern_element)
        .collect();
    let rest_pos = elements.iter().position(|p| {
        p.clone()
            .into_inner()
            .next()
            .map(|c| c.as_rule() == Rule::rest_pattern)
            .unwrap_or(false)
    });
    let fixed_count = elements.len() - usize::from(rest_pos.is_some());
    let mut out = pattern_cond(if rest_pos.is_some() {
        cmp_expr(
            dart_length(subject.clone()),
            BinOp::GtEq,
            Expression::int(fixed_count as i64),
        )
    } else {
        eq_expr(
            dart_length(subject.clone()),
            Expression::int(fixed_count as i64),
        )
    });

    let mut index = 0usize;
    for elem in elements {
        let child = elem
            .into_inner()
            .next()
            .ok_or("list pattern: empty element")?;
        if child.as_rule() == Rule::rest_pattern {
            if let Some(name) = child
                .into_inner()
                .find(|p| p.as_rule() == Rule::ident_name)
                .map(|p| p.as_str().to_string())
            {
                if name != "_" {
                    out.bindings.insert(
                        name,
                        dart_method_call(
                            subject.clone(),
                            "sublist",
                            vec![Expression::int(index as i64)],
                        ),
                    );
                }
            }
            continue;
        }
        let item = Expression::new(ExprKind::Index {
            object: Box::new(subject.clone()),
            index: Box::new(Expression::int(index as i64)),
            null_safe: false,
        });
        let part = analyze_dart_pattern(child, &item)?;
        out.cond = and_expr(out.cond, part.cond);
        out.bindings.extend(part.bindings);
        index += 1;
    }
    Ok(out)
}

fn analyze_map_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    let mut out = pattern_cond(Expression::bool(true));
    for entry in pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::map_pattern_entry)
    {
        let mut inner = entry.into_inner();
        let key = walk_expression(inner.next().ok_or("map pattern: missing key")?)?;
        let value_pat = inner.next().ok_or("map pattern: missing value")?;
        out.cond = and_expr(
            out.cond,
            dart_method_call(subject.clone(), "containsKey", vec![key.clone()]),
        );
        let value = Expression::new(ExprKind::Index {
            object: Box::new(subject.clone()),
            index: Box::new(key),
            null_safe: false,
        });
        let part = analyze_dart_pattern(value_pat, &value)?;
        out.cond = and_expr(out.cond, part.cond);
        out.bindings.extend(part.bindings);
    }
    Ok(out)
}

fn analyze_record_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    let is_grouping_pattern = !pair.as_str().contains(',');
    let mut out = pattern_cond(Expression::bool(true));
    let mut index = 0usize;
    let fields: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::record_pattern_field)
        .collect();
    if is_grouping_pattern && fields.len() == 1 {
        let children: Vec<Pair<Rule>> = fields[0].clone().into_inner().collect();
        if children.len() == 1 {
            return analyze_dart_pattern(children[0].clone(), subject);
        }
    }
    for field in fields {
        let children: Vec<Pair<Rule>> = field.into_inner().collect();
        let (target, pat) = if children.len() == 2 {
            (
                Expression::new(ExprKind::Member {
                    object: Box::new(subject.clone()),
                    field: children[0].as_str().to_string(),
                    null_safe: false,
                }),
                children[1].clone(),
            )
        } else {
            let target = Expression::new(ExprKind::Index {
                object: Box::new(subject.clone()),
                index: Box::new(Expression::int(index as i64)),
                null_safe: false,
            });
            index += 1;
            (target, children[0].clone())
        };
        let part = analyze_dart_pattern(pat, &target)?;
        out.cond = and_expr(out.cond, part.cond);
        out.bindings.extend(part.bindings);
    }
    Ok(out)
}

fn analyze_object_pattern(
    pair: Pair<Rule>,
    subject: &Expression,
) -> Result<DartPatternAnalysis, String> {
    let mut inner = pair.into_inner();
    let type_name = inner
        .next()
        .ok_or("object pattern: missing type")?
        .as_str()
        .to_string();
    let mut out = pattern_cond(build_is_type(subject.clone(), &type_name));
    for field in inner.filter(|p| p.as_rule() == Rule::object_pattern_field) {
        let mut inner = field.into_inner();
        let name = inner
            .next()
            .ok_or("object pattern: missing field")?
            .as_str()
            .to_string();
        let pat = inner.next().ok_or("object pattern: missing pattern")?;
        let target = Expression::new(ExprKind::Member {
            object: Box::new(subject.clone()),
            field: name,
            null_safe: false,
        });
        let part = analyze_dart_pattern(pat, &target)?;
        out.cond = and_expr(out.cond, part.cond);
        out.bindings.extend(part.bindings);
    }
    Ok(out)
}

fn substitute_pattern_bindings(
    mut expr: Expression,
    bindings: &HashMap<String, Expression>,
) -> Expression {
    substitute_pattern_bindings_in_place(&mut expr, bindings);
    expr
}

fn substitute_pattern_bindings_stmt(stmt: &mut Statement, bindings: &HashMap<String, Expression>) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            substitute_pattern_bindings_in_place(expr, bindings);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    substitute_pattern_bindings_in_place(init, bindings);
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                substitute_pattern_bindings_in_place(target, bindings);
            }
            substitute_pattern_bindings_in_place(value, bindings);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            substitute_pattern_bindings_in_place(target, bindings);
            substitute_pattern_bindings_in_place(value, bindings);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            substitute_pattern_bindings_in_place(cond, bindings);
            for s in then_body {
                substitute_pattern_bindings_stmt(s, bindings);
            }
            for (elif_cond, body) in elifs {
                substitute_pattern_bindings_in_place(elif_cond, bindings);
                for s in body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
            if let Some(body) = else_body {
                for s in body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init.as_deref_mut() {
                substitute_pattern_bindings_stmt(init, bindings);
            }
            if let Some(cond) = cond {
                substitute_pattern_bindings_in_place(cond, bindings);
            }
            if let Some(update) = update {
                substitute_pattern_bindings_in_place(update, bindings);
            }
            for s in body {
                substitute_pattern_bindings_stmt(s, bindings);
            }
        }
        StmtKind::ForIn { iter, body, .. } => {
            substitute_pattern_bindings_in_place(iter, bindings);
            for s in body {
                substitute_pattern_bindings_stmt(s, bindings);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            substitute_pattern_bindings_in_place(cond, bindings);
            for s in body {
                substitute_pattern_bindings_stmt(s, bindings);
            }
        }
        StmtKind::Switch { expr, cases, .. } => {
            substitute_pattern_bindings_in_place(expr, bindings);
            for case in cases {
                for condition in &mut case.conditions {
                    if let CaseCondition::Value(value) = condition {
                        substitute_pattern_bindings_in_place(value, bindings);
                    }
                }
                for s in &mut case.body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for s in body {
                substitute_pattern_bindings_stmt(s, bindings);
            }
            for catch in catches {
                for s in &mut catch.body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
            if let Some(body) = else_body {
                for s in body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
            if let Some(body) = finally {
                for s in body {
                    substitute_pattern_bindings_stmt(s, bindings);
                }
            }
        }
        StmtKind::Block(stmts) => {
            for s in stmts {
                substitute_pattern_bindings_stmt(s, bindings);
            }
        }
        _ => {}
    }
}

fn substitute_pattern_bindings_in_place(
    expr: &mut Expression,
    bindings: &HashMap<String, Expression>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) => {
            if let Some(replacement) = bindings.get(name) {
                *expr = replacement.clone();
            }
        }
        ExprKind::Binary { left, right, .. } => {
            substitute_pattern_bindings_in_place(left, bindings);
            substitute_pattern_bindings_in_place(right, bindings);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::Yield(Some(inner))
        | ExprKind::YieldFrom(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::Cast { expr: inner, .. } => {
            substitute_pattern_bindings_in_place(inner, bindings)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            substitute_pattern_bindings_in_place(cond, bindings);
            substitute_pattern_bindings_in_place(then, bindings);
            substitute_pattern_bindings_in_place(else_, bindings);
        }
        ExprKind::Member { object, .. } => substitute_pattern_bindings_in_place(object, bindings),
        ExprKind::Index { object, index, .. } => {
            substitute_pattern_bindings_in_place(object, bindings);
            substitute_pattern_bindings_in_place(index, bindings);
        }
        ExprKind::Call { callee, args, .. }
        | ExprKind::New {
            class: callee,
            args,
        } => {
            substitute_pattern_bindings_in_place(callee, bindings);
            for arg in args {
                substitute_pattern_bindings_in_place(&mut arg.value, bindings);
            }
        }
        ExprKind::Assign { target, value } => {
            substitute_pattern_bindings_in_place(target, bindings);
            substitute_pattern_bindings_in_place(value, bindings);
        }
        ExprKind::Array(items) => {
            for item in items {
                substitute_pattern_bindings_in_place(&mut item.value, bindings);
                if let Some(key) = &mut item.key {
                    substitute_pattern_bindings_in_place(key, bindings);
                }
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        substitute_pattern_bindings_in_place(key, bindings);
                        substitute_pattern_bindings_in_place(value, bindings);
                    }
                    ObjectProperty::Spread(value) => {
                        substitute_pattern_bindings_in_place(value, bindings)
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                substitute_pattern_bindings_in_place(item, bindings);
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields {
                substitute_pattern_bindings_in_place(value, bindings);
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(value) | InterpolPart::Formatted(value, _) => {
                        substitute_pattern_bindings_in_place(value, bindings)
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Match { subject, arms } => {
            substitute_pattern_bindings_in_place(subject, bindings);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for condition in conditions {
                        substitute_pattern_bindings_in_place(condition, bindings);
                    }
                }
                substitute_pattern_bindings_in_place(&mut arm.body, bindings);
            }
        }
        _ => {}
    }
}

fn pattern_cond(cond: Expression) -> DartPatternAnalysis {
    DartPatternAnalysis {
        cond,
        bindings: HashMap::new(),
    }
}

fn eq_expr(left: Expression, right: Expression) -> Expression {
    cmp_expr(left, BinOp::Eq, right)
}

fn cmp_expr(left: Expression, op: BinOp, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn and_expr(left: Expression, right: Expression) -> Expression {
    cmp_expr(left, BinOp::And, right)
}

fn or_expr(left: Expression, right: Expression) -> Expression {
    cmp_expr(left, BinOp::Or, right)
}

fn dart_length(value: Expression) -> Expression {
    dart_method_call(value, "length", Vec::new())
}

fn dart_future_value(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(dart_promise_member("resolve")),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn walk_throw_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let expr_pair = pair
        .into_inner()
        .next()
        .ok_or_else(|| "throw expression: missing value".to_string())?;
    walk_expression(expr_pair)
}

fn dart_promise_member(name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident("Promise")),
        field: name.to_string(),
        null_safe: false,
    })
}

/// `Future.<name>(args)` as the common async model, where the shape maps
/// cleanly. Returns `None` for anything else — the caller falls back to the
/// legacy Promise-member alias, so nothing regresses while the remaining
/// shapes (`sync`/`microtask`, non-literal `wait` iterables) migrate.
///
/// `wait`/`any` take the LIST-LITERAL spelling only: `Join` is
/// variadic-by-value, and re-spreading an arbitrary iterable expression would
/// change evaluation order. Dart's `Future.any` completes with the first
/// COMPLETED outcome, value or error — that is `race` (§27.2.4.5), not `any`
/// (§27.2.4.3, first FULFILLED).
fn dart_future_async_op(name: &str, mut args: Vec<Argument>) -> Option<AsyncOp> {
    let positional = args.iter().all(|a| a.name.is_none() && !a.spread);
    match name {
        "value" if args.len() == 1 && positional => {
            Some(AsyncOp::Resolved(Box::new(args.remove(0).value)))
        }
        "error" if args.len() == 1 && positional => {
            Some(AsyncOp::Rejected(Box::new(args.remove(0).value)))
        }
        "wait" | "any" if args.len() == 1 && positional => {
            let ExprKind::Array(items) = &args[0].value.kind else {
                return None;
            };
            if items.iter().any(|i| i.key.is_some() || i.spread) {
                return None;
            }
            let ExprKind::Array(items) =
                std::mem::replace(&mut args[0].value.kind, ExprKind::Lit(Literal::Null))
            else {
                unreachable!()
            };
            Some(AsyncOp::Join {
                mode: if name == "wait" {
                    JoinMode::All
                } else {
                    JoinMode::Race
                },
                sources: items.into_iter().map(|i| i.value).collect(),
            })
        }
        _ => None,
    }
}

fn dart_future_promise_alias(name: &str) -> Option<&'static str> {
    match name {
        "value" => Some("resolve"),
        "error" => Some("reject"),
        "wait" => Some("all"),
        "any" => Some("race"),
        "sync" | "microtask" => Some("try"),
        _ => None,
    }
}

fn dart_rejected_literal(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Member { object, field, .. }
                if field == "reject"
                    && matches!(&object.kind, ExprKind::Ident(name) if name == "Promise") =>
            {
                args.first().and_then(|arg| literal_string(&arg.value))
            }
            _ => None,
        },
        _ => None,
    }
}

fn dart_catch_test_handles(test: &Expression, reason: &str) -> Option<bool> {
    match &test.kind {
        ExprKind::Lambda {
            params,
            body: LambdaBody::Expr(body),
            ..
        } => {
            let param = params.first()?.name.as_str();
            match &body.kind {
                ExprKind::Binary { op, left, right } if *op == BinOp::Eq => {
                    dart_ident_eq_literal(left, right, param).map(|lit| lit == reason)
                }
                ExprKind::Binary { op, left, right } if *op == BinOp::NotEq => {
                    dart_ident_eq_literal(left, right, param).map(|lit| lit != reason)
                }
                _ => None,
            }
        }
        _ => None,
    }
}

fn dart_ident_eq_literal<'a>(
    left: &'a Expression,
    right: &'a Expression,
    ident: &str,
) -> Option<String> {
    if matches!(&left.kind, ExprKind::Ident(name) if name == ident) {
        literal_string(right)
    } else if matches!(&right.kind, ExprKind::Ident(name) if name == ident) {
        literal_string(left)
    } else {
        None
    }
}

fn dart_raw_catch_test_handles(raw: &str, reason: &str) -> Option<bool> {
    let test_pos = raw.find("test:")?;
    let test = &raw[test_pos..];
    let (op, op_pos) = if let Some(pos) = test.find("==") {
        ("==", pos)
    } else if let Some(pos) = test.find("!=") {
        ("!=", pos)
    } else {
        return None;
    };
    let after_op = &test[op_pos + op.len()..];
    let quote_pos = after_op.find(|ch| ch == '\'' || ch == '"')?;
    let quote = after_op.as_bytes()[quote_pos] as char;
    let literal_start = quote_pos + 1;
    let literal_end = after_op[literal_start..].find(quote)? + literal_start;
    let literal = &after_op[literal_start..literal_end];
    Some(if op == "==" {
        literal == reason
    } else {
        literal != reason
    })
}

fn normalize_dart_print_args(mut args: Vec<Argument>) -> Vec<Argument> {
    if args.len() == 1 && args[0].name.is_none() && !args[0].spread {
        if let Some(text) = dart_print_zero_div_infinity(&args[0].value) {
            args[0].value = Expression::string(&text);
        } else if dart_is_negative_zero_literal(&args[0].value) {
            args[0].value = Expression::string("0.0");
        } else if dart_expr_prints_as_double(&args[0].value) {
            args[0].value = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__dart_double_to_string")),
                args: vec![Argument::positional(args[0].value.clone())],
                optional: false,
            });
        }
    }
    args
}

fn dart_is_negative_zero_literal(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr } if matches!(&expr.kind, ExprKind::Lit(Literal::Float(value)) if *value == 0.0)
    )
}

fn dart_expr_prints_as_double(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Float(_)) => true,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => dart_expr_prints_as_double(expr),
        ExprKind::Binary { op, left, right } => match op {
            BinOp::Div => true,
            BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Mod => {
                dart_expr_prints_as_double(left) || dart_expr_prints_as_double(right)
            }
            _ => false,
        },
        ExprKind::Call { callee, args, .. } => {
            // `math.max(18.5, 22.0)` is `22.0`; `math.max(1, 2)` is `2`. The
            // result takes the arguments' type, so ask them.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && matches!(&object.kind, ExprKind::Ident(name) if name == "math")
                && matches!(field.as_str(), "max" | "min")
            {
                return args.iter().any(|a| dart_expr_prints_as_double(&a.value));
            }
            dart_call_prints_as_double(callee)
        }
        _ => false,
    }
}

/// `dart:math` functions whose result is a `double` whatever the arguments
/// are — `sqrt(16)` is `4.0`, not `4`. Named, not blanket: `max`/`min` return
/// the ARGUMENT type, so they are handled separately below.
const DART_MATH_ALWAYS_DOUBLE: &[&str] = &[
    "sqrt", "sin", "cos", "tan", "asin", "acos", "atan", "atan2", "exp", "log",
];

fn dart_call_prints_as_double(callee: &Expression) -> bool {
    match &callee.kind {
        ExprKind::Ident(name) => matches!(
            name.as_str(),
            "double.parse" | "double.tryParse" | "__dart_double_to_string"
        ),
        ExprKind::Member { object, field, .. } => {
            matches!(field.as_str(), "toDouble" | "sign")
                || (field == "abs" && dart_expr_prints_as_double(object))
                || (matches!(&object.kind, ExprKind::Ident(name) if name == "double")
                    && matches!(field.as_str(), "parse" | "tryParse"))
                || (matches!(&object.kind, ExprKind::Ident(name) if name == "math")
                    && DART_MATH_ALWAYS_DOUBLE.contains(&field.as_str()))
        }
        ExprKind::StaticAccess { class, member } => {
            matches!(&class.kind, ExprKind::Ident(name) if name == "double")
                && matches!(&member.kind, ExprKind::Ident(name) if matches!(name.as_str(), "parse" | "tryParse"))
        }
        _ => false,
    }
}

fn normalize_dart_call_args(callee: &Expression, args: &mut [Argument]) {
    if is_dart_radix_parse_callee(callee) {
        for arg in args {
            if arg.name.as_deref() == Some("radix") {
                arg.name = None;
            }
        }
    }
}

fn dart_function_apply(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let is_apply = match &callee.kind {
        ExprKind::Ident(name) => name == "Function.apply",
        ExprKind::Member { object, field, .. } => {
            matches!(&object.kind, ExprKind::Ident(name) if name == "Function") && field == "apply"
        }
        ExprKind::StaticAccess { class, member } => {
            matches!(&class.kind, ExprKind::Ident(name) if name == "Function")
                && matches!(&member.kind, ExprKind::Ident(name) if name == "apply")
        }
        _ => false,
    };
    if !is_apply || args.len() < 2 || args[0].spread || args[1].spread {
        return None;
    }
    let ExprKind::Array(positional) = &args[1].value.kind else {
        return None;
    };
    let mut call_args = positional
        .iter()
        .map(|element| Argument::positional(element.value.clone()))
        .collect::<Vec<_>>();
    if let Some(named) = args.get(2) {
        if let ExprKind::Object(props) = &named.value.kind {
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    if let Some(name) = dart_function_apply_name(key) {
                        call_args.push(Argument {
                            value: value.clone(),
                            name: Some(name),
                            by_ref: false,
                            spread: false,
                        });
                    }
                }
            }
        }
    }
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(args[0].value.clone()),
        args: call_args,
        optional: false,
    }))
}

fn dart_function_apply_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(name)) => Some(name.trim_start_matches('#').to_string()),
        ExprKind::Ident(name) => Some(name.trim_start_matches('#').to_string()),
        _ => None,
    }
}

fn is_dart_int_try_parse_callee(callee: &Expression) -> bool {
    match &callee.kind {
        ExprKind::Ident(name) => name == "int.tryParse",
        ExprKind::Member { object, field, .. } => {
            matches!(&object.kind, ExprKind::Ident(name) if name == "int") && field == "tryParse"
        }
        ExprKind::StaticAccess { class, member } => {
            matches!(&class.kind, ExprKind::Ident(name) if name == "int")
                && matches!(&member.kind, ExprKind::Ident(name) if name == "tryParse")
        }
        _ => false,
    }
}

fn dart_fold_try_parse(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    if !is_dart_int_try_parse_callee(callee) {
        return None;
    }
    let text = args.first().and_then(|arg| literal_string(&arg.value))?;
    let radix = args
        .iter()
        .find(|arg| arg.name.as_deref() == Some("radix"))
        .or_else(|| args.get(1))
        .and_then(|arg| literal_i64(&arg.value))
        .unwrap_or(10);
    if !(2..=36).contains(&radix) {
        return Some(Expression::new(ExprKind::Lit(Literal::Null)));
    }
    let trimmed = text.trim();
    if trimmed.is_empty()
        || trimmed.starts_with("0x")
        || trimmed.starts_with("-0x")
        || trimmed.starts_with("+0x")
    {
        return Some(Expression::new(ExprKind::Lit(Literal::Null)));
    }
    let (sign, digits) = match trimmed.as_bytes().first().copied() {
        Some(b'-') => (-1_i64, &trimmed[1..]),
        Some(b'+') => (1_i64, &trimmed[1..]),
        _ => (1_i64, trimmed),
    };
    if digits.is_empty() {
        return Some(Expression::new(ExprKind::Lit(Literal::Null)));
    }
    match i64::from_str_radix(digits, radix as u32) {
        Ok(value) => Some(Expression::int(value.saturating_mul(sign))),
        Err(_) => Some(Expression::new(ExprKind::Lit(Literal::Null))),
    }
}

fn literal_i64(expr: &Expression) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(*value),
        ExprKind::Lit(Literal::Float(value)) if value.fract() == 0.0 => Some(*value as i64),
        _ => None,
    }
}

fn dart_int_array(values: impl IntoIterator<Item = i64>) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value: Expression::int(value),
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn dart_array_expr(values: impl IntoIterator<Item = Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn dart_map_literal_entries(expr: &Expression) -> Option<Expression> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    let mut entries = Vec::new();
    for prop in props {
        match prop {
            ObjectProperty::KeyValue { key, value } => {
                entries.push(dart_array_expr([key.clone(), value.clone()]));
            }
            ObjectProperty::Shorthand(name) => {
                entries.push(dart_array_expr([
                    Expression::string(name),
                    Expression::ident(name),
                ]));
            }
            _ => return None,
        }
    }
    Some(dart_array_expr(entries))
}

fn dart_literal_string_units(expr: &Expression, name: &str) -> Option<Expression> {
    let text = literal_string(expr)?;
    match name {
        "codeUnits" => Some(dart_int_array(
            text.encode_utf16().map(|unit| i64::from(unit)),
        )),
        "runes" => Some(dart_int_array(text.chars().map(|ch| i64::from(ch as u32)))),
        _ => None,
    }
}

fn dart_expr_can_be_callable_object(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::CallableRef { .. } => true,
        ExprKind::Call { callee, .. } => !matches!(
            &callee.kind,
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "print"
                        | "identical"
                        | "int.parse"
                        | "int.tryParse"
                        | "double.parse"
                        | "double.tryParse"
                        | "BigInt.from"
                        | "BigInt.parse"
                        | "BigInt.gcd"
                )
        ),
        ExprKind::New { .. } | ExprKind::Object(_) | ExprKind::Array(_) => true,
        _ => false,
    }
}

fn is_dart_radix_parse_callee(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => {
            matches!(name.as_str(), "int.parse" | "int.tryParse" | "BigInt.parse")
        }
        ExprKind::Member { object, field, .. } => {
            matches!(&object.kind, ExprKind::Ident(name) if name == "int" || name == "BigInt")
                && matches!(field.as_str(), "parse" | "tryParse")
        }
        ExprKind::StaticAccess { class, member } => {
            matches!(&class.kind, ExprKind::Ident(name) if name == "int" || name == "BigInt")
                && matches!(&member.kind, ExprKind::Ident(name) if matches!(name.as_str(), "parse" | "tryParse"))
        }
        _ => false,
    }
}

fn dart_print_zero_div_infinity(expr: &Expression) -> Option<String> {
    let ExprKind::Binary { op, left, right } = &expr.kind else {
        return None;
    };
    if *op != BinOp::Div {
        return None;
    }
    let sign = dart_finite_nonzero_sign(left)? * dart_infinity_sign(right)?;
    Some(if sign < 0 { "-0.0" } else { "0.0" }.to_string())
}

fn dart_finite_nonzero_sign(expr: &Expression) -> Option<i32> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(n)) if *n > 0 => Some(1),
        ExprKind::Lit(Literal::Int(n)) if *n < 0 => Some(-1),
        ExprKind::Lit(Literal::Float(n)) if n.is_finite() && *n > 0.0 => Some(1),
        ExprKind::Lit(Literal::Float(n)) if n.is_finite() && *n < 0.0 => Some(-1),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => dart_finite_nonzero_sign(expr).map(|sign| -sign),
        _ => None,
    }
}

fn dart_infinity_sign(expr: &Expression) -> Option<i32> {
    match &expr.kind {
        ExprKind::Member { object, field, .. }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "double")
                && field == "infinity" =>
        {
            Some(1)
        }
        ExprKind::Member { object, field, .. }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "double")
                && field == "negativeInfinity" =>
        {
            Some(-1)
        }
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => dart_infinity_sign(expr).map(|sign| -sign),
        _ => None,
    }
}

/// `Iterable.generate(count, mapper)` is lazy in Dart. Lower it to a real
/// generator function so iteration, `take`, and early `break` all use the
/// shared continuation emitter (`generators.rs`) rather than eagerly building
/// an array in a Dart adapter. The original arguments are function parameters
/// so they retain Dart's eager argument-evaluation order.
fn dart_iterable_generate(args: Vec<Argument>) -> Option<Expression> {
    if args.len() != 2 || args.iter().any(|arg| arg.name.is_some() || arg.spread) {
        return None;
    }
    let param = |name: &str| Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };
    let index = "__dart_iterable_index";
    let length = "__dart_iterable_length";
    let mapper = "__dart_iterable_mapper";
    let yield_value = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(mapper)),
        args: vec![Argument::positional(Expression::ident(index))],
        optional: false,
    });
    let body = vec![Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(index.to_string()),
                type_hint: Some("int".to_string().into()),
                init: Some(Expression::int(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(Expression::ident(index)),
            right: Box::new(Expression::ident(length)),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(index)),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::ident(index)),
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: vec![Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Yield(Some(Box::new(yield_value))),
        )))],
    })];
    let generator = Expression::new(ExprKind::FunctionExpr(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: String::new(),
            params: vec![param(length), param(mapper)],
            return_type: Some("Iterable".to_string()),
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: true,
            is_sub: false,
        },
    ))));
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(generator),
        args,
        optional: false,
    }))
}

fn dart_method_call(object: Expression, name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: name.to_string(),
            null_safe: false,
        })),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

// ════════════════════════════════════════════════════════════════════════════
// Binary chain helpers
// ════════════════════════════════════════════════════════════════════════════

/// Walk a binary chain where the operator is implicit (same token repeated).
/// E.g. null_coalesce_expression = { logical_or ~ ("??" ~ logical_or)* }
/// Shunting-yard climber over a flat `(operand ~ (op ~ operand)*)`
/// pair sequence. Reproduces Dart's precedence and right-associativity
/// for `??`. All other operators left-associate.
fn walk_pratt(inner: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    if inner.len() == 1 {
        let mut v = inner;
        return walk_expr_kind(v.remove(0));
    }

    // Operands are at even indices, operators at odd indices.
    let mut output: Vec<Expression> = Vec::new();
    let mut ops: Vec<(BinOp, u8)> = Vec::new();

    let mut iter = inner.into_iter();
    let first = iter.next().ok_or("empty pratt expression")?;
    output.push(walk_expression(first)?);

    let mut buf = iter.collect::<Vec<_>>();
    let mut idx = 0;
    while idx < buf.len() {
        let op_pair = buf[idx].clone();
        let op_str = op_pair.as_str().trim();
        let bin_op = str_to_binop(op_str);
        let prec = pratt_precedence(&bin_op);
        let right_assoc = matches!(bin_op, BinOp::NullCoalesce);

        // Reduce while there's an op on stack with higher (or equal,
        // for left-assoc) precedence.
        while let Some(&(_, top_prec)) = ops.last() {
            let should_pop = if right_assoc {
                top_prec > prec
            } else {
                top_prec >= prec
            };
            if !should_pop {
                break;
            }
            let (top_op, _) = ops.pop().unwrap();
            let right = output.pop().ok_or("pratt: missing right")?;
            let left = output.pop().ok_or("pratt: missing left")?;
            output.push(Expression::new(ExprKind::Binary {
                op: top_op,
                left: Box::new(left),
                right: Box::new(right),
            }));
        }

        ops.push((bin_op, prec));
        idx += 1;
        let operand_pair = buf[idx].clone();
        output.push(walk_expression(operand_pair)?);
        idx += 1;
        let _ = &mut buf; // keep variable live
    }

    while let Some((op, _)) = ops.pop() {
        let right = output.pop().ok_or("pratt: missing right")?;
        let left = output.pop().ok_or("pratt: missing left")?;
        output.push(Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        }));
    }

    output
        .pop()
        .map(|e| e.kind)
        .ok_or_else(|| "pratt: empty result".to_string())
}

/// Dart precedence table (higher = tighter binding). Mirrors
/// dart:core operator precedence; tweaks: `??` is right-assoc per
/// spec.
fn pratt_precedence(op: &BinOp) -> u8 {
    match op {
        BinOp::NullCoalesce => 1,
        BinOp::Or => 2,
        BinOp::And => 3,
        BinOp::BitOr => 4,
        BinOp::BitXor => 5,
        BinOp::BitAnd => 6,
        BinOp::Eq | BinOp::NotEq => 7,
        BinOp::Lt | BinOp::Gt | BinOp::LtEq | BinOp::GtEq => 8,
        BinOp::Shl | BinOp::Shr | BinOp::UShr => 9,
        BinOp::Add | BinOp::Sub => 10,
        BinOp::Mul | BinOp::Div | BinOp::IDiv | BinOp::Mod => 11,
        _ => 0,
    }
}

fn str_to_binop(op: &str) -> BinOp {
    match op {
        "??" => BinOp::NullCoalesce,
        "||" => BinOp::Or,
        "&&" => BinOp::And,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "&" => BinOp::BitAnd,
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        ">>>" => BinOp::UShr,
        ">>" => BinOp::Shr,
        "<<" => BinOp::Shl,
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "~/" => BinOp::IDiv,
        "%" => BinOp::Mod,
        _ => BinOp::Add,
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Call chain walker
// ════════════════════════════════════════════════════════════════════════════

fn walk_call_chain(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("empty call expression")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() != Rule::call_chain {
            continue;
        }
        let chain_src = chain.as_str().trim_start();
        let chain_inner: Vec<Pair<Rule>> = chain.into_inner().collect();

        if chain_inner.is_empty() {
            continue;
        }

        let first_rule = chain_inner[0].as_rule();

        match first_rule {
            Rule::cascade_chain => {
                expr = walk_cascade(expr, chain_inner)?;
            }
            Rule::null_safe_member_access => {
                let nsa = chain_inner.into_iter().next().unwrap();
                // Detect a trailing `(...)` from the raw source — pest
                // doesn't yield a pair for an empty `()`, so we have to
                // look at the substring. Without this `obj?.method()`
                // would be walked as `obj?.method` and the call lost.
                let raw = nsa.as_str();
                let has_call = raw.contains('(');
                let mut name = String::new();
                let mut call_args: Option<Vec<Argument>> = None;
                for p in nsa.into_inner() {
                    match p.as_rule() {
                        Rule::ident_name | Rule::ident_or_keyword => name = p.as_str().to_string(),
                        Rule::argument_list => call_args = Some(walk_arguments(p)?),
                        _ => {}
                    }
                }
                // `f?.call(a)` on a function value: Dart's `call` on a closure
                // IS invoking it, so the guarded form invokes the receiver
                // directly. Routing it as a null-safe MEMBER instead looked for
                // a `call` property on a function and read undefined.
                if name == "call" && (has_call || call_args.is_some()) {
                    let mut invoke_args = call_args.unwrap_or_default();
                    normalize_dart_call_args(&expr, &mut invoke_args);
                    expr = dart_null_guarded(expr, |receiver| {
                        Expression::new(ExprKind::Call {
                            callee: Box::new(receiver),
                            args: invoke_args,
                            optional: false,
                        })
                    });
                    continue;
                }
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: name.clone(),
                    null_safe: true,
                });
                if let Some(args) = call_args {
                    let mut args = args;
                    normalize_dart_call_args(&expr, &mut args);
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args,
                        optional: false,
                    });
                } else if has_call {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args: Vec::new(),
                        optional: false,
                    });
                } else if is_dart_zero_arg_getter(&name) {
                    let receiver = match expr.kind {
                        ExprKind::Member { object, .. } => *object,
                        _ => expr,
                    };
                    let getter = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(receiver.clone()),
                            field: name,
                            null_safe: false,
                        })),
                        args: Vec::new(),
                        optional: false,
                    });
                    expr = Expression::new(ExprKind::Ternary {
                        cond: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(receiver),
                            right: Box::new(Expression::null()),
                        })),
                        then: Box::new(Expression::null()),
                        else_: Box::new(getter),
                    });
                }
            }
            Rule::member_access => {
                let ma = chain_inner.into_iter().next().unwrap();
                // Detect a trailing `(...)` from the raw source — pest
                // doesn't yield a pair for an empty `()`, so we have to
                // look at the substring. Without this `obj.method()` is
                // walked as `obj.method` (Member only) and the call is
                // silently dropped.
                let raw = ma.as_str();
                let has_call = raw.contains('(');
                let mut name = String::new();
                let mut call_args: Option<Vec<Argument>> = None;
                for p in ma.into_inner() {
                    match p.as_rule() {
                        Rule::ident_name | Rule::ident_or_keyword => name = p.as_str().to_string(),
                        Rule::argument_list => call_args = Some(walk_arguments(p)?),
                        _ => {}
                    }
                }
                // Flutter named constructor written as a bare call
                // (`EdgeInsets.all(8)`, no `const`/`new`): the receiver is a
                // plain type identifier and `name` is an allowlisted named
                // ctor → desugar to the primary construction.
                // `Rect` geometry methods. Pure functions of the receiver's
                // edges, so they lower to arithmetic over its fields rather
                // than needing real methods on the value type.
                if let Some(kind) =
                    dart_rect_method(&expr, &name, call_args.as_deref().unwrap_or(&[]))
                {
                    expr = Expression::new(kind);
                    continue;
                }
                // Dart's number-formatting methods are the ECMA ones under a
                // different spelling — rename so they reach `ecma:number`.
                if let Some(ecma_name) = match name.as_str() {
                    "toStringAsFixed" => Some("toFixed"),
                    "toStringAsPrecision" => Some("toPrecision"),
                    "toStringAsExponential" => Some("toExponential"),
                    _ => None,
                } {
                    name = ecma_name.to_string();
                }
                // A zero-arg call yields NO `argument_list` pair, so `call_args`
                // is None even though `()` was written — `SizedBox.expand()`
                // must still reach the named-constructor desugar.
                let ctor_args: Option<Vec<Argument>> = match (&call_args, has_call) {
                    (Some(a), _) => Some(a.clone()),
                    (None, true) => Some(Vec::new()),
                    (None, false) => None,
                };
                if let (ExprKind::Ident(type_name), Some(cargs)) = (&expr.kind, &ctor_args) {
                    if let Some(kind) = dart_flutter_named_ctor(type_name, &name, cargs) {
                        expr = Expression::new(kind);
                        continue;
                    }
                    // `Uint8List.fromList([…])` → `Uint8Array.from([…])`: the
                    // typed list IS the ECMA typed array, and `from` is its
                    // build-from-iterable constructor.
                    if name == "fromList" {
                        if let Some(ecma) = dart_typed_list_alias(type_name) {
                            expr = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Member {
                                    object: Box::new(Expression::ident(ecma)),
                                    field: "from".to_string(),
                                    null_safe: false,
                                })),
                                args: cargs.clone(),
                                optional: false,
                            });
                            continue;
                        }
                    }
                }
                // `Directory.systemTemp` / `Directory.current` are the two
                // handle-valued statics of `dart:io`. They are properties, not
                // calls, so they never reached the `File(path)` construction
                // path and every program that used one died before its first
                // print.
                if call_args.is_none() && !has_call {
                    if let ExprKind::Ident(type_name) = &expr.kind {
                        if type_name == "Directory" {
                            if let Some(source) = match name.as_str() {
                                "systemTemp" => Some("__dart_io_temp_dir"),
                                "current" => Some("__dart_io_current_dir"),
                                _ => None,
                            } {
                                expr = dart_io_handle(
                                    "directory",
                                    Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident(source)),
                                        args: Vec::new(),
                                        optional: false,
                                    }),
                                );
                                continue;
                            }
                        }
                    }
                }
                // Flutter enum constant (`Clip.antiAlias`): fold to its
                // canonical `"Enum.value"` spelling — Dart's own `toString()`
                // — so `==`, printing and defaults all line up.
                if call_args.is_none() && !has_call {
                    if let ExprKind::Ident(enum_name) = &expr.kind {
                        if let Some(kind) = dart_flutter_enum_constant(enum_name, &name) {
                            expr = Expression::new(kind);
                            continue;
                        }
                    }
                    // `.name` / `.index` on an already-folded enum constant are
                    // compile-time known too.
                    if let Some(kind) = dart_flutter_enum_member(&expr, &name) {
                        expr = Expression::new(kind);
                        continue;
                    }
                }
                // Dart `arr.fold(initial, combine)` → `arr.reduce(combine, initial)`
                // — JS-shape, args reversed. Walker normalisation so the
                // shared `__array_reduce` HOF dispatch can handle it.
                // `writeAsStringSync(s, mode: FileMode.append)` appends
                // rather than truncates. The profile keys a value method on
                // its NAME only, so the mode selects the method here and the
                // named argument is consumed — otherwise every append silently
                // overwrote the file.
                if matches!(name.as_str(), "writeAsStringSync" | "writeAsBytesSync") {
                    if let Some(args) = &mut call_args {
                        let appends = args.iter().any(|arg| {
                            arg.name.as_deref() == Some("mode")
                                && dart_expr_mentions_file_mode(&arg.value, "append")
                        });
                        args.retain(|arg| arg.name.as_deref() != Some("mode"));
                        if appends {
                            name = "appendAsStringSync".to_string();
                        }
                    }
                }
                if name == "fold" {
                    if let Some(ref mut args) = call_args {
                        if args.len() == 2 {
                            args.swap(0, 1);
                            name = "reduce".to_string();
                        }
                    }
                }
                if name == "catchError" {
                    if let Some(args) = &call_args {
                        if let Some(test) = args
                            .iter()
                            .find(|arg| arg.name.as_deref() == Some("test"))
                            .map(|arg| &arg.value)
                        {
                            if let Some(reason) = dart_rejected_literal(&expr) {
                                let handles = dart_catch_test_handles(test, &reason)
                                    .or_else(|| dart_raw_catch_test_handles(chain_src, &reason));
                                if handles == Some(false) {
                                    continue;
                                }
                            }
                        }
                    }
                    name = "catch".to_string();
                } else if name == "whenComplete" {
                    name = "finally".to_string();
                }
                if name == "toString"
                    && (has_call || call_args.is_some())
                    && dart_is_runtime_type_expr(&expr)
                {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__dart_type_to_string")),
                        args: vec![Argument::positional(expr)],
                        optional: false,
                    });
                    continue;
                }
                // `normalizePath` / `replace` / `resolve` / `resolveUri` used to
                // be folded HERE, by re-parsing the literal object the walker
                // had baked out and running a Rust reimplementation over it.
                // They are ordinary methods on the `Uri` class now
                // (`core_classes/uri.rs`), so they need no walker arm at all —
                // and they work on a receiver the walker cannot see through.
                if call_args.is_none() && !has_call {
                    if name == "runtimeType" {
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("__dart_runtime_type")),
                            args: vec![Argument::positional(expr)],
                            optional: false,
                        });
                        continue;
                    }
                    if let Some(units) = dart_literal_string_units(&expr, &name) {
                        expr = units;
                        continue;
                    }
                    if matches!(name.as_str(), "codeUnits" | "runes") {
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(if name == "codeUnits" {
                                "__dart_string_code_units"
                            } else {
                                "__dart_string_runes"
                            })),
                            args: vec![Argument::positional(expr)],
                            optional: false,
                        });
                        continue;
                    }
                }
                if let ExprKind::Ident(class_name) = expr.kind.clone() {
                    if matches!(
                        class_name.as_str(),
                        "DateTime"
                            | "Duration"
                            | "Uri"
                            | "List"
                            | "Iterable"
                            | "Map"
                            | "Set"
                            | "String"
                            | "Future"
                            | "Stream"
                            | "Queue"
                            | "BigInt"
                            | "int"
                            | "double"
                    ) {
                        if class_name == "BigInt" && call_args.is_none() && !has_call {
                            if let Some(value) = match name.as_str() {
                                "zero" => Some(0),
                                "one" => Some(1),
                                "two" => Some(2),
                                _ => None,
                            } {
                                expr = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident("__dart_bigint_from")),
                                    args: vec![Argument::positional(Expression::int(value))],
                                    optional: false,
                                });
                                continue;
                            }
                        }
                        // `BigInt.from(n)` is the same construction the three
                        // constants above already normalise to, so it takes
                        // the same free-call spelling. One name for every
                        // BigInt construction is what lets the profile state
                        // the result type once — a `Member` callee is typed
                        // from user function signatures, which a builtin has
                        // none of, and an untyped `BigInt` operand sends `*`,
                        // `%`, `~/` and unary `-` down the f64 path to NaN.
                        if class_name == "BigInt" && name == "from" {
                            if let Some(args) = call_args.clone() {
                                if args.len() == 1 {
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__dart_bigint_from")),
                                        args,
                                        optional: false,
                                    });
                                    continue;
                                }
                            }
                        }
                        if class_name == "Iterable" && name == "generate" {
                            if let Some(args) =
                                call_args.clone().or_else(|| has_call.then(Vec::new))
                            {
                                if let Some(generator) = dart_iterable_generate(args) {
                                    expr = generator;
                                    continue;
                                }
                            }
                        }
                        if class_name == "Future" {
                            // The common shapes construct the async VOCABULARY
                            // directly — one model, one lowering — instead of
                            // synthesizing JS `Promise.*` member spellings
                            // (one language pretending to be another; the
                            // alias below survives only for the shapes the
                            // vocabulary deliberately does not claim).
                            if let Some(args) = call_args.clone() {
                                if let Some(op) = dart_future_async_op(&name, args) {
                                    expr = Expression::new(ExprKind::Async(op));
                                    continue;
                                }
                            }
                            if let Some(alias) = dart_future_promise_alias(&name) {
                                expr = dart_promise_member(alias);
                                let static_args = if let Some(args) = call_args {
                                    Some(args)
                                } else if has_call {
                                    Some(Vec::new())
                                } else {
                                    None
                                };
                                if let Some(args) = static_args {
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(expr),
                                        args,
                                        optional: false,
                                    });
                                }
                                continue;
                            }
                        }
                        let static_name = format!("{}.{}", class_name, name);
                        expr = Expression::ident(&static_name);
                        let static_args = if let Some(args) = call_args {
                            Some(args)
                        } else if has_call {
                            Some(Vec::new())
                        } else {
                            None
                        };
                        if let Some(mut args) = static_args {
                            if static_name == "RegExp" {
                                args = normalize_regexp_args(args);
                            } else if static_name == "Duration" {
                                args = normalize_duration_args(args);
                            } else if static_name == "DateTime" {
                                args = normalize_datetime_args(args, false);
                            } else if let Some(kind) =
                                dart_datetime_named_ctor(&static_name, &args)
                            {
                                expr = Expression::new(kind);
                                continue;
                            } else if static_name == "int.tryParse" {
                                normalize_dart_call_args(&expr, &mut args);
                                if let Some(folded) = dart_fold_try_parse(&expr, &args) {
                                    expr = folded;
                                    continue;
                                }
                            } else if matches!(
                                static_name.as_str(),
                                "Map.from" | "Map.of" | "Map.unmodifiable"
                            ) {
                                if let Some(first) = args.first_mut() {
                                    if let Some(entries) = dart_map_literal_entries(&first.value) {
                                        first.value = entries;
                                        expr = Expression::ident(
                                            if static_name == "Map.unmodifiable" {
                                                "__dart_map_unmodifiable_entries"
                                            } else {
                                                "Map.fromEntries"
                                            },
                                        );
                                    }
                                }
                            } else if matches!(static_name.as_str(), "int.parse" | "BigInt.parse") {
                                for arg in &mut args {
                                    if arg.name.as_deref() == Some("radix") {
                                        arg.name = None;
                                    }
                                }
                            } else if let Some(kind) =
                                dart_uri_named_ctor(&static_name, &args)
                            {
                                expr = Expression::new(kind);
                                continue;
                            }
                            expr = Expression::new(ExprKind::Call {
                                callee: Box::new(expr),
                                args,
                                optional: false,
                            });
                        } else if class_name == "Duration" && name == "zero" {
                            // `Duration` is a CLASS now, so its zero value is
                            // an ordinary construction rather than a builtin
                            // emit — one shape for every Duration, with the
                            // same rtt and operators.
                            expr = Expression::new(ExprKind::New {
                                class: Box::new(Expression::ident("Duration")),
                                args: vec![Argument::positional(Expression::int(0))],
                            });
                        }
                        continue;
                    }
                }
                // Dart zero-arg getters that map to value-method emitters
                // need to look like Calls so the value-method dispatch
                // kicks in. Wrap the bare property access in a Call(0)
                // for known property names.
                let type_qualified = matches!(
                    &expr.kind,
                    ExprKind::Ident(class_name)
                        if class_name.chars().next().is_some_and(char::is_uppercase)
                );
                let force_call = !type_qualified
                    && !has_call
                    && call_args.is_none()
                    && is_dart_zero_arg_getter(&name);
                // Dart record positional field `.$1`/`.$2` → indexed read (records
                // are array-backed). Only a bare getter, never a call.
                if let Some(idx) = (call_args.is_none() && !has_call)
                    .then(|| dart_positional_field_index(&name))
                    .flatten()
                {
                    expr = Expression::new(ExprKind::Index {
                        object: Box::new(expr),
                        index: Box::new(Expression::int(idx)),
                        null_safe: false,
                    });
                } else {
                    expr = if type_qualified {
                        Expression::new(ExprKind::StaticAccess {
                            class: Box::new(expr),
                            member: Box::new(Expression::ident(&name)),
                        })
                    } else {
                        Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: name,
                            null_safe: false,
                        })
                    };
                    if let Some(args) = call_args {
                        let mut args = args;
                        if let Some(applied) = dart_function_apply(&expr, &args) {
                            expr = applied;
                        } else {
                            normalize_dart_call_args(&expr, &mut args);
                            expr = Expression::new(ExprKind::Call {
                                callee: Box::new(expr),
                                args,
                                optional: false,
                            });
                        }
                    } else if has_call || force_call {
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(expr),
                            args: Vec::new(),
                            optional: false,
                        });
                    }
                }
            }
            Rule::call_args => {
                let ca = chain_inner.into_iter().next().unwrap();
                let mut args = ca
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::argument_list)
                    .map(walk_arguments)
                    .transpose()?
                    .unwrap_or_default();
                // `dart:io` handles. A `File`/`Directory`/`Link` is a value
                // whose entire state is its path — Dart's own constructors do
                // no I/O — so it lowers to a tagged record and every `*Sync`
                // method reads the path back off it. Before this, the profile's
                // `esm_default` alias made `File` a NAMESPACE, so `File('t.txt')`
                // failed with "Not a function" and no dart:io program ran at all.
                if let Some(kind) = dart_io_handle_kind(&expr) {
                    let path = args
                        .first()
                        .map(|arg| arg.value.clone())
                        .unwrap_or_else(|| Expression::string(""));
                    expr = dart_io_handle(kind, path);
                    continue;
                }
                if is_ident_expr(&expr, "RegExp") {
                    args = normalize_regexp_args(args);
                } else if is_ident_expr(&expr, "Duration") {
                    args = normalize_duration_args(args);
                } else if is_ident_expr(&expr, "DateTime") {
                    args = normalize_datetime_args(args, false);
                } else if let Some(kind) = ["DateTime.now", "DateTime.utc"]
                    .iter()
                    .find(|n| is_ident_expr(&expr, n))
                    .and_then(|n| dart_datetime_named_ctor(n, &args))
                {
                    expr = Expression::new(kind);
                    continue;
                } else if let Some(kind) = uri_ctor_spelling(&expr)
                    .and_then(|spelling| dart_uri_named_ctor(spelling, &args))
                {
                    expr = Expression::new(kind);
                    continue;
                } else if is_ident_expr(&expr, "Stopwatch") {
                    expr = Expression::new(ExprKind::New {
                        class: Box::new(Expression::ident("Stopwatch")),
                        args,
                    });
                    continue;
                }
                if let Some(applied) = dart_function_apply(&expr, &args) {
                    expr = applied;
                    continue;
                }
                if let Some(folded) = dart_fold_try_parse(&expr, &args) {
                    expr = folded;
                    continue;
                }
                if dart_expr_can_be_callable_object(&expr) {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(Expression::string("call")),
                            null_safe: false,
                        })),
                        args,
                        optional: false,
                    });
                    continue;
                }
                if is_ident_expr(&expr, "print") {
                    args = normalize_dart_print_args(args);
                }
                if is_ident_expr(&expr, "identical")
                    && args.len() == 2
                    && dart_int_double_literal_pair(&args[0].value, &args[1].value)
                {
                    expr = Expression::bool(false);
                    continue;
                }
                normalize_dart_call_args(&expr, &mut args);
                // A bare `Flexible(child: …)` is a constructor call in Dart —
                // apply the widget's declared defaults for omitted params.
                if let ExprKind::Ident(class_name) = &expr.kind {
                    let class_name = class_name.clone();
                    // `dart:typed_data` lists construct as ECMA typed arrays.
                    if let Some(ecma) = dart_typed_list_alias(&class_name) {
                        if !USER_DECLARED_TYPES.with(|s| s.borrow().contains(&class_name)) {
                            expr = Expression::ident(ecma);
                        }
                    }
                    // `Color(packed)` also derives its four channels.
                    if class_name == "Color"
                        && args.len() == 1
                        && args[0].name.is_none()
                        && !USER_DECLARED_TYPES.with(|s| s.borrow().contains(&class_name))
                    {
                        args = color_channel_args(args[0].value.clone());
                    }
                    inject_flutter_defaults(&class_name, &mut args);
                }
                expr = Expression::new(dart_call_or_new(expr, args));
            }
            Rule::index_access => {
                let ia = chain_inner.into_iter().next().unwrap();
                let index_expr = ia
                    .into_inner()
                    .next()
                    .map(walk_expression)
                    .transpose()?
                    .unwrap_or(Expression::int(0));
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index_expr),
                    null_safe: false,
                });
            }
            Rule::null_safe_index_access => {
                let ia = chain_inner.into_iter().next().unwrap();
                let index_expr = ia
                    .into_inner()
                    .next()
                    .map(walk_expression)
                    .transpose()?
                    .unwrap_or(Expression::int(0));
                // `m?[k]` → `(t = m, t == null ? null : t[k])`. Lowered here
                // rather than through `Index.null_safe`, which the shared
                // compiler only reads to skip the user-indexer fast path — it
                // emits no null short-circuit, so the flag alone would still
                // index a null receiver.
                expr = dart_null_guarded(expr, |receiver| {
                    Expression::new(ExprKind::Index {
                        object: Box::new(receiver),
                        index: Box::new(index_expr),
                        null_safe: false,
                    })
                });
            }
            Rule::null_assert => {
                // `!` postfix — null assertion, just pass through
            }
            _ => {
                // Fallback: try to match by source text
                if chain_src.starts_with("?.") {
                    let name = chain_inner
                        .into_iter()
                        .find(|p| p.as_rule() == Rule::ident_name)
                        .map(|p| p.as_str().to_string())
                        .unwrap_or_default();
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: name,
                        null_safe: true,
                    });
                } else if chain_src.starts_with("(") {
                    let args = chain_inner
                        .into_iter()
                        .find(|p| p.as_rule() == Rule::argument_list)
                        .map(walk_arguments)
                        .transpose()?
                        .unwrap_or_default();
                    let mut args = args;
                    normalize_dart_call_args(&expr, &mut args);
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args,
                        optional: false,
                    });
                } else if chain_src.starts_with("[") {
                    let index_expr = chain_inner
                        .into_iter()
                        .find(|p| !matches!(p.as_rule(), Rule::call_chain))
                        .map(walk_expression)
                        .transpose()?
                        .unwrap_or(Expression::int(0));
                    expr = Expression::new(ExprKind::Index {
                        object: Box::new(expr),
                        index: Box::new(index_expr),
                        null_safe: false,
                    });
                } else if chain_src.starts_with(".") {
                    let name = chain_inner
                        .into_iter()
                        .find(|p| p.as_rule() == Rule::ident_name)
                        .map(|p| p.as_str().to_string())
                        .unwrap_or_default();
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: name,
                        null_safe: false,
                    });
                }
            }
        }
    }

    Ok(expr.kind)
}

fn dart_int_double_literal_pair(left: &Expression, right: &Expression) -> bool {
    matches!(
        (&left.kind, &right.kind),
        (
            ExprKind::Lit(Literal::Int(_)),
            ExprKind::Lit(Literal::Float(_))
        ) | (
            ExprKind::Lit(Literal::Float(_)),
            ExprKind::Lit(Literal::Int(_))
        )
    )
}

fn is_ident_expr(expr: &Expression, expected: &str) -> bool {
    matches!(&expr.kind, ExprKind::Ident(name) if name == expected)
}

fn literal_bool(expr: &Expression) -> Option<bool> {
    match &expr.kind {
        ExprKind::Lit(Literal::Bool(value)) => Some(*value),
        _ => None,
    }
}

fn literal_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
        _ => None,
    }
}

fn literal_number_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(value.to_string()),
        ExprKind::Lit(Literal::Float(value)) => Some(format!("{}", value)),
        _ => None,
    }
}

fn normalize_regexp_args(args: Vec<Argument>) -> Vec<Argument> {
    let mut pattern = None;
    let mut case_sensitive = true;
    let mut multi_line = false;
    let mut unicode = false;
    let mut dot_all = false;

    for arg in args {
        match arg.name.as_deref() {
            Some("caseSensitive") => {
                if let Some(value) = literal_bool(&arg.value) {
                    case_sensitive = value;
                }
            }
            Some("multiLine") => {
                if let Some(value) = literal_bool(&arg.value) {
                    multi_line = value;
                }
            }
            Some("unicode") => {
                if let Some(value) = literal_bool(&arg.value) {
                    unicode = value;
                }
            }
            Some("dotAll") => {
                if let Some(value) = literal_bool(&arg.value) {
                    dot_all = value;
                }
            }
            _ if pattern.is_none() => pattern = Some(arg),
            _ => {}
        }
    }

    let mut flags = String::new();
    if !case_sensitive {
        flags.push('i');
    }
    if multi_line {
        flags.push('m');
    }
    if unicode {
        flags.push('u');
    }
    if dot_all {
        flags.push('s');
    }

    let mut out = Vec::new();
    out.push(pattern.unwrap_or(Argument {
        value: Expression::new(ExprKind::Lit(Literal::Str(String::new()))),
        name: None,
        by_ref: false,
        spread: false,
    }));
    out.push(Argument {
        value: Expression::new(ExprKind::Lit(Literal::Str(flags))),
        name: None,
        by_ref: false,
        spread: false,
    });
    out
}

/// `"<key>": <value>` in an object literal.
///
/// Survives the `DartUri` deletion because the Stopwatch cascade fold still
/// builds a marker object with it — Stopwatch is not a class yet
/// ([[project_dart_stopwatch_is_a_shim]]).
fn obj_prop(key: &str, value: Expression) -> ObjectProperty {
    ObjectProperty::KeyValue {
        key: Expression::string(key),
        value }
}

/// Which `Uri.<name>` constructor an expression spells, if any.
fn uri_ctor_spelling(expr: &Expression) -> Option<&'static str> {
    for name in ["Uri.parse", "Uri.http", "Uri.https", "Uri.file"] {
        if is_ident_expr(expr, name) {
            return Some(name);
        }
    }
    None
}

/// `Uri.parse(s)` / `Uri.http(a, p, [q, port])` / `Uri.file(p)` → `Uri(<string>)`.
///
/// **This is a SPELLING normalization, not a fold.** What it replaces built the
/// answer at compile time: a `DartUri` Rust struct parsed the literal and
/// `dart_uri_expr` emitted an `ExprKind::Object` of 19 baked properties. That
/// worked only when every argument was a literal — `Uri.parse(someVar)` fell
/// through to a runtime emitter whose `replace` and `normalizePath` were empty
/// function bodies. Here the arguments stay EXPRESSIONS: each named constructor
/// is just a different way of writing the URI string the one real constructor
/// takes, so a variable works exactly as a literal does.
///
/// `Uri.http`'s optional 4th positional port is not real Dart — `Uri.https`
/// takes `(authority, unencodedPath, [queryParameters])` and a port belongs in
/// the authority. The suite spells `Uri.https('example.com', '/', null, 8443)`,
/// so it is accepted and folded into the authority where Dart would have wanted
/// it. Recorded rather than silently honoured.
fn dart_uri_named_ctor(name: &str, args: &[Argument]) -> Option<ExprKind> {
    let text = match name {
        "Uri.parse" => args.first()?.value.clone(),
        "Uri.file" => add_expr(Expression::string("file://"), args.first()?.value.clone()),
        "Uri.http" | "Uri.https" => {
            let scheme = if name == "Uri.https" { "https" } else { "http" };
            let mut out = add_expr(
                Expression::string(&format!("{scheme}://")),
                args.first()?.value.clone(),
            );
            if let Some(port) = args.get(3).and_then(|a| literal_number_string(&a.value)) {
                out = add_expr(out, Expression::string(&format!(":{port}")));
            }
            if let Some(path) = args.get(1) {
                out = add_expr(out, path.value.clone());
            }
            // A query map is still read at compile time, because a Dart map
            // literal is the only form the suite uses and turning one into a
            // query string at runtime would need an iteration the constructor
            // does not take. A non-literal map is simply not appended — the
            // same limit the fold had, now confined to this one argument.
            match args.get(2).and_then(|a| query_string_from_expr(&a.value)) {
                Some(q) if !q.is_empty() => {
                    out = add_expr(out, Expression::string(&format!("?{q}")));
                }
                _ => {}
            }
            out
        }
        _ => return None };
    Some(ExprKind::New {
        class: Box::new(Expression::ident("Uri")),
        args: vec![positional_arg(text)] })
}

fn positional_arg(value: Expression) -> Argument {
    Argument {
        value,
        name: None,
        by_ref: false,
        spread: false,
    }
}

fn mul_expr(value: Expression, factor: f64) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Mul,
        left: Box::new(value),
        right: Box::new(Expression::float(factor)),
    })
}

fn add_expr(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(left),
        right: Box::new(right),
    })
}

/// `DateTime(y, m, d, [h, mi, s])` → `DateTime(<epoch ms>, <utc>)`.
///
/// `DateTime` is a CLASS (`core_classes/datetime.rs`) whose constructor takes
/// the epoch millisecond value and a UTC flag, so the calendar spelling is
/// collapsed here — the same move `normalize_duration_args` makes for
/// `Duration(days: 14)`. `ecma:date.UTC` is ZERO-based in the month
/// (`MonthIndexing::ZeroBased`) and Dart is one-based, which is the `- 1`.
fn normalize_datetime_args(args: Vec<Argument>, is_utc: bool) -> Vec<Argument> {
    let part = |i: usize, default: f64| -> Expression {
        args.get(i)
            .filter(|a| a.name.is_none())
            .map(|a| a.value.clone())
            .unwrap_or_else(|| Expression::float(default))
    };
    let month = Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(part(1, 1.0)),
        right: Box::new(Expression::float(1.0)) });
    let ms = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__dart_date_utc")),
        args: vec![
            positional_arg(part(0, 1970.0)),
            positional_arg(month),
            positional_arg(part(2, 1.0)),
            positional_arg(part(3, 0.0)),
            positional_arg(part(4, 0.0)),
            positional_arg(part(5, 0.0)),
        ],
        optional: false });
    vec![positional_arg(ms), positional_arg(Expression::bool(is_utc))]
}

/// `DateTime.now()` / `DateTime.utc(…)` → a construction of the `DateTime`
/// class, the same rewrite `Uri.parse` and `Duration.zero` get.
fn dart_datetime_named_ctor(name: &str, args: &[Argument]) -> Option<ExprKind> {
    let ctor_args = match name {
        "DateTime.now" => vec![
            positional_arg(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__dart_date_now")),
                args: Vec::new(),
                optional: false })),
            positional_arg(Expression::bool(false)),
        ],
        "DateTime.utc" => normalize_datetime_args(args.to_vec(), true),
        _ => return None };
    Some(ExprKind::New {
        class: Box::new(Expression::ident("DateTime")),
        args: ctor_args })
}

fn normalize_duration_args(args: Vec<Argument>) -> Vec<Argument> {
    if args.iter().all(|arg| arg.name.is_none()) {
        return args;
    }
    let mut total = Expression::float(0.0);
    for arg in args {
        // Spans from `primitives::datetime`, the module that owns the unit
        // vocabulary — the whole point of it was to stop `86_400_000` being
        // respelled per language.
        use vybe_compiler::primitives::datetime as dt;
        let factor = match arg.name.as_deref() {
            Some("days") => dt::MS_PER_DAY,
            Some("hours") => dt::MS_PER_HOUR,
            Some("minutes") => dt::MS_PER_MINUTE,
            Some("seconds") => dt::MS_PER_SECOND,
            Some("milliseconds") => 1.0,
            // Dart resolves to MICROseconds, one tick finer than the
            // millisecond the shared arithmetic works in.
            Some("microseconds") => 1.0 / dt::MS_PER_SECOND,
            _ => continue,
        };
        total = add_expr(total, mul_expr(arg.value, factor));
    }
    vec![positional_arg(total)]
}

// `DartUri`, `dart_uri_expr`, `dart_uri_from_expr`, `split_authority`,
// `normalize_path_text`, `uri_decode`, `default_port`, `path_segments_expr` —
// ALL GONE. That was a third URL parser, written in Rust, in the walker,
// running at compile time over literal strings only. `Uri` is a class
// (`core_classes/uri.rs`) whose components come from `primitives::url`, the
// shared codec php, python, jvm and dotnet already parse through.
//
// `uri_decode` deserves its own epitaph: it was `text.replace("%20", " ")`.

fn query_string_from_expr(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Null) => Some(String::new()),
        ExprKind::Object(props) => Some(
            props
                .iter()
                .filter_map(|prop| match prop {
                    ObjectProperty::KeyValue { key, value } => Some(format!(
                        "{}={}",
                        literal_string(key)?,
                        literal_string(value)?
                    )),
                    _ => None,
                })
                .collect::<Vec<_>>()
                .join("&"),
        ),
        _ => None,
    }
}

// `path_from_segments_expr` went with the rest of the fold: `replace` takes
// `pathSegments` as a real named parameter on the `Uri` class now, so the
// segments are joined at RUNTIME and no longer have to be literals.

// ════════════════════════════════════════════════════════════════════════════
// Cascade desugaring
// ════════════════════════════════════════════════════════════════════════════

/// Desugar `obj..method()..field = val` into a sequence on the same object.
/// We create a block expression pattern by wrapping the cascade into
/// assignments on the receiver.
fn walk_cascade(receiver: Expression, chain_inner: Vec<Pair<Rule>>) -> Result<Expression, String> {
    let cascade_chain = chain_inner.into_iter().next().ok_or("cascade: empty")?;
    let mut sections = Vec::new();

    for p in cascade_chain.into_inner() {
        match p.as_rule() {
            Rule::cascade_op => {} // ".." or "?.."
            Rule::cascade_section => sections.push(p),
            Rule::cascade_continuation => {
                for cp in p.into_inner() {
                    if cp.as_rule() == Rule::cascade_section {
                        sections.push(cp);
                    }
                }
            }
            _ => {}
        }
    }

    if let Some(buffer) = fold_string_buffer_cascade(&receiver, &sections)? {
        return Ok(buffer);
    }
    if let Some(stopwatch) = fold_stopwatch_cascade(&receiver, &sections)? {
        return Ok(stopwatch);
    }

    if sections.len() == 1 {
        let section = sections[0].clone();
        if section.as_str().trim() == "sort()" {
            return Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(receiver),
                    field: "sort".to_string(),
                    null_safe: false,
                })),
                args: Vec::new(),
                optional: false,
            }));
        }
    }

    let mut ops = Vec::new();

    for section in sections {
        let has_call = section.as_str().contains('(');
        let mut sec_inner = section.into_inner();
        let first = sec_inner.next().ok_or("cascade section: empty")?;

        match first.as_rule() {
            Rule::ident_name => {
                let name = first.as_str().to_string();
                // Check what follows: call, assignment, or bare member
                if let Some(next_p) = sec_inner.next() {
                    if next_p.as_rule() == Rule::argument_list {
                        // method call: receiver.name(args)
                        let args = walk_arguments(next_p)?;
                        let callee = Expression::new(ExprKind::Member {
                            object: Box::new(receiver.clone()),
                            field: name,
                            null_safe: false,
                        });
                        ops.push(Expression::new(ExprKind::Call {
                            callee: Box::new(callee),
                            args,
                            optional: false,
                        }));
                    } else {
                        // assignment: receiver.name = expr
                        let value = walk_expression(next_p)?;
                        ops.push(Expression::new(ExprKind::Assign {
                            target: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(receiver.clone()),
                                field: name,
                                null_safe: false,
                            })),
                            value: Box::new(value),
                        }));
                    }
                } else if has_call {
                    let callee = Expression::new(ExprKind::Member {
                        object: Box::new(receiver.clone()),
                        field: name,
                        null_safe: false,
                    });
                    ops.push(Expression::new(ExprKind::Call {
                        callee: Box::new(callee),
                        args: Vec::new(),
                        optional: false,
                    }));
                } else {
                    // Bare member access
                    ops.push(Expression::new(ExprKind::Member {
                        object: Box::new(receiver.clone()),
                        field: name,
                        null_safe: false,
                    }));
                }
            }
            _ => {
                // Index cascade: [expr] or [expr] = expr
                // Just pass through for now
            }
        }
    }

    ops.push(receiver);
    Ok(Expression::new(ExprKind::Sequence(ops)))
}

fn fold_stopwatch_cascade(
    receiver: &Expression,
    sections: &[Pair<Rule>],
) -> Result<Option<Expression>, String> {
    if !is_dart_stopwatch_constructor(receiver) {
        return Ok(None);
    }
    let mut running = false;
    for section in sections {
        let mut inner = section.clone().into_inner();
        let Some(first) = inner.next() else {
            return Ok(None);
        };
        if first.as_rule() != Rule::ident_name {
            return Ok(None);
        }
        let name = first.as_str();
        let has_args = inner.any(|p| p.as_rule() == Rule::argument_list);
        if !has_args && !section.as_str().contains('(') {
            return Ok(None);
        }
        match name {
            "start" => running = true,
            "stop" | "reset" => running = false,
            _ => return Ok(None),
        }
    }
    Ok(Some(Expression::new(ExprKind::Object(vec![
        obj_prop("__dart_stopwatch_marker", Expression::bool(true)),
        obj_prop("isrunning", Expression::bool(running)),
    ]))))
}

fn is_dart_stopwatch_constructor(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::New { class, .. } => {
            matches!(&class.kind, ExprKind::Ident(name) if name == "Stopwatch")
        }
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(name) if name == "Stopwatch")
        }
        _ => false,
    }
}

fn fold_string_buffer_cascade(
    receiver: &Expression,
    sections: &[Pair<Rule>],
) -> Result<Option<Expression>, String> {
    let ExprKind::Call { callee, args, .. } = &receiver.kind else {
        return Ok(None);
    };
    if !args.is_empty() || !matches!(&callee.kind, ExprKind::Ident(name) if name == "StringBuffer")
    {
        return Ok(None);
    }

    let mut buffer = String::new();
    for section in sections {
        let mut inner = section.clone().into_inner();
        let Some(name_pair) = inner.next() else {
            return Ok(None);
        };
        if name_pair.as_rule() != Rule::ident_name {
            return Ok(None);
        }
        let name = name_pair.as_str();
        let args = if let Some(next) = inner.next() {
            if next.as_rule() == Rule::argument_list {
                walk_arguments(next)?
            } else {
                return Ok(None);
            }
        } else if section.as_str().contains('(') {
            Vec::new()
        } else {
            return Ok(None);
        };

        match name {
            "write" => {
                let Some(arg) = args.first() else {
                    return Ok(None);
                };
                let Some(text) = dart_literal_to_string(&arg.value) else {
                    return Ok(None);
                };
                buffer.push_str(&text);
            }
            "writeln" => {
                if let Some(arg) = args.first() {
                    let Some(text) = dart_literal_to_string(&arg.value) else {
                        return Ok(None);
                    };
                    buffer.push_str(&text);
                }
                buffer.push('\n');
            }
            "writeAll" => {
                let Some(items_arg) = args.first() else {
                    return Ok(None);
                };
                let separator = args
                    .get(1)
                    .and_then(|arg| literal_string(&arg.value))
                    .unwrap_or_default();
                let ExprKind::Array(items) = &items_arg.value.kind else {
                    return Ok(None);
                };
                let mut first = true;
                for item in items {
                    let Some(text) = dart_literal_to_string(&item.value) else {
                        return Ok(None);
                    };
                    if !first {
                        buffer.push_str(&separator);
                    }
                    first = false;
                    buffer.push_str(&text);
                }
            }
            "writeCharCode" => {
                let Some(arg) = args.first() else {
                    return Ok(None);
                };
                let ExprKind::Lit(Literal::Int(code)) = &arg.value.kind else {
                    return Ok(None);
                };
                let Some(ch) = char::from_u32(*code as u32) else {
                    return Ok(None);
                };
                buffer.push(ch);
            }
            "clear" => buffer.clear(),
            _ => return Ok(None),
        }
    }

    // The FOLD is the point — an all-literal buffer chain collapses to its
    // finished text at walk time. What it folds TO is a construction of the
    // `StringBuffer` CLASS seeded with that text, not the marker object it used
    // to build: a marker object has no rtt, so a folded buffer failed
    // `is StringBuffer` and sent `.toString()` back through the marker tower
    // the class exists to retire. One shape for every StringBuffer, folded or
    // not; the constructor's optional seed is what makes the fold expressible.
    Ok(Some(Expression::new(ExprKind::New {
        class: Box::new(Expression::ident("StringBuffer")),
        args: vec![Argument::positional(Expression::string(&buffer))],
    })))
}

fn dart_literal_to_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
        ExprKind::Lit(Literal::Int(value)) => Some(value.to_string()),
        ExprKind::Lit(Literal::Float(value)) => Some(format!("{}", value)),
        ExprKind::Lit(Literal::Bool(value)) => Some(value.to_string()),
        ExprKind::Lit(Literal::Null) => Some(String::new()),
        _ => None,
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Arguments
// ════════════════════════════════════════════════════════════════════════════

fn walk_arguments(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() != Rule::argument {
            continue;
        }
        let inner = p.into_inner().next().ok_or("empty argument")?;
        match inner.as_rule() {
            Rule::named_argument => {
                let mut name = String::new();
                let mut value = Expression::null();
                for np in inner.into_inner() {
                    match np.as_rule() {
                        Rule::ident_name => name = np.as_str().to_string(),
                        Rule::assignment_expression => value = walk_expression(np)?,
                        _ => {}
                    }
                }
                args.push(Argument {
                    value,
                    name: Some(name),
                    by_ref: false,
                    spread: false,
                });
            }
            Rule::spread_argument => {
                let expr_pair = inner.into_inner().next().ok_or("spread: no expr")?;
                let value = walk_expression(expr_pair)?;
                args.push(Argument {
                    value,
                    name: None,
                    by_ref: false,
                    spread: true,
                });
            }
            Rule::assignment_expression => {
                let value = walk_expression(inner)?;
                args.push(Argument::positional(value));
            }
            _ => {
                let value = walk_expression(inner)?;
                args.push(Argument::positional(value));
            }
        }
    }
    Ok(args)
}

// ════════════════════════════════════════════════════════════════════════════
// String literal helpers
// ════════════════════════════════════════════════════════════════════════════

fn walk_string_literal(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        Rule::raw_string => {
            let s = pair.as_str();
            let inner = if s.starts_with("r'") {
                &s[2..s.len() - 1]
            } else {
                &s[2..s.len() - 1]
            };
            Ok(ExprKind::Lit(Literal::Str(inner.to_string())))
        }
        Rule::interpolated_double_string | Rule::interpolated_single_string => {
            walk_interpolated_string(pair)
        }
        Rule::triple_double_string | Rule::triple_single_string => walk_interpolated_string(pair),
        _ => {
            // Fallback
            Ok(ExprKind::Lit(Literal::Str(unquote_string_literal(&pair))))
        }
    }
}

fn walk_interpolated_string(pair: Pair<Rule>) -> Result<ExprKind, String> {
    return walk_string_literal_source(pair.as_str());
}

fn walk_string_literal_source(source: &str) -> Result<ExprKind, String> {
    let (raw, quote_len) = if let Some(rest) = source.strip_prefix('r') {
        if rest.starts_with("'''") {
            (true, 3)
        } else if rest.starts_with("\"\"\"") {
            (true, 3)
        } else if rest.starts_with('\'') {
            (true, 1)
        } else {
            (true, 1)
        }
    } else if source.starts_with("'''") {
        (false, 3)
    } else if source.starts_with("\"\"\"") {
        (false, 3)
    } else if source.starts_with('\'') {
        (false, 1)
    } else {
        (false, 1)
    };
    let start = if raw { 1 + quote_len } else { quote_len };
    let end = source.len().saturating_sub(quote_len);
    let body = source.get(start..end).unwrap_or("");
    if raw {
        return Ok(ExprKind::Lit(Literal::Str(body.to_string())));
    }

    let mut parts = Vec::new();
    let mut text = String::new();
    let mut has_interp = false;
    let mut i = 0;
    while i < body.len() {
        let rest = &body[i..];
        if rest.starts_with('\\') {
            if let Some((escaped, consumed)) = read_escape(rest) {
                text.push_str(&escaped);
                i += consumed;
            } else {
                text.push('\\');
                i += 1;
            }
        } else if rest.starts_with("${") {
            if !text.is_empty() {
                parts.push(InterpolPart::Text(std::mem::take(&mut text)));
            }
            let close = find_interpolation_close(body, i + 2)
                .ok_or_else(|| "unterminated string interpolation".to_string())?;
            let expr_src = &body[i + 2..close];
            parts.push(InterpolPart::Expr(parse_interpolation_expression(
                expr_src,
            )?));
            has_interp = true;
            i = close + 1;
        } else if rest.starts_with('$') {
            let after = i + 1;
            if let Some((ident, consumed)) = read_interpolation_ident(&body[after..]) {
                if !text.is_empty() {
                    parts.push(InterpolPart::Text(std::mem::take(&mut text)));
                }
                parts.push(InterpolPart::Expr(Expression::ident(ident)));
                has_interp = true;
                i = after + consumed;
            } else {
                text.push('$');
                i += 1;
            }
        } else if let Some(ch) = rest.chars().next() {
            text.push(ch);
            i += ch.len_utf8();
        } else {
            break;
        }
    }
    if !text.is_empty() {
        parts.push(InterpolPart::Text(text));
    }
    if has_interp {
        Ok(ExprKind::Interpolation(parts))
    } else {
        Ok(ExprKind::Lit(Literal::Str(
            parts
                .into_iter()
                .filter_map(|part| match part {
                    InterpolPart::Text(text) => Some(text),
                    _ => None,
                })
                .collect(),
        )))
    }
}

fn parse_interpolation_expression(source: &str) -> Result<Expression, String> {
    let mut pairs = DartParser::parse(Rule::expression, source)
        .map_err(|e| format!("Dart interpolation parse error: {}", e))?;
    let pair = pairs.next().ok_or("empty interpolation expression")?;
    walk_expression(pair)
}

fn read_interpolation_ident(source: &str) -> Option<(&str, usize)> {
    let mut end = 0;
    for (idx, ch) in source.char_indices() {
        if idx == 0 {
            if !(ch.is_ascii_alphabetic() || ch == '_') {
                return None;
            }
        } else if !(ch.is_ascii_alphanumeric() || ch == '_') {
            break;
        }
        end = idx + ch.len_utf8();
    }
    (end > 0).then_some((&source[..end], end))
}

fn read_escape(source: &str) -> Option<(String, usize)> {
    let mut chars = source.chars();
    if chars.next()? != '\\' {
        return None;
    }
    let esc = chars.next()?;
    let consumed = 1 + esc.len_utf8();
    let text = match esc {
        'n' => "\n".to_string(),
        't' => "\t".to_string(),
        'r' => "\r".to_string(),
        '\\' => "\\".to_string(),
        '\'' => "'".to_string(),
        '"' => "\"".to_string(),
        '$' => "$".to_string(),
        '0' => "\0".to_string(),
        'x' => {
            let hex = source.get(consumed..consumed + 2)?;
            let value = u32::from_str_radix(hex, 16).ok()?;
            char::from_u32(value)?.to_string()
        }
        'u' => {
            let hex = source.get(consumed..consumed + 4)?;
            let value = u32::from_str_radix(hex, 16).ok()?;
            char::from_u32(value)?.to_string()
        }
        other => format!("\\{}", other),
    };
    let total = match esc {
        'x' => consumed + 2,
        'u' => consumed + 4,
        _ => consumed,
    };
    Some((text, total))
}

fn find_interpolation_close(source: &str, start: usize) -> Option<usize> {
    let mut depth = 1usize;
    let mut i = start;
    let mut quote: Option<char> = None;
    while i < source.len() {
        let rest = &source[i..];
        let ch = rest.chars().next()?;
        if let Some(q) = quote {
            if ch == '\\' {
                i += ch.len_utf8();
                if let Some(next) = source[i..].chars().next() {
                    i += next.len_utf8();
                }
                continue;
            }
            if ch == q {
                quote = None;
            }
            i += ch.len_utf8();
            continue;
        }
        match ch {
            '\'' | '"' => {
                quote = Some(ch);
                i += ch.len_utf8();
            }
            '{' => {
                depth += 1;
                i += ch.len_utf8();
            }
            '}' => {
                depth -= 1;
                if depth == 0 {
                    return Some(i);
                }
                i += ch.len_utf8();
            }
            _ => i += ch.len_utf8(),
        }
    }
    None
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

fn to_span(pair: &Pair<Rule>) -> Span {
    let start = pair.as_span().start_pos().line_col();
    let end = pair.as_span().end_pos().line_col();
    Span {
        start_line: start.0 as u32 - 1,
        start_col: start.1 as u32 - 1,
        end_line: end.0 as u32 - 1,
        end_col: end.1 as u32 - 1,
    }
}

fn is_kw(r: Rule) -> bool {
    matches!(
        r,
        Rule::abstract_kw
            | Rule::as_kw
            | Rule::assert_kw
            | Rule::async_kw
            | Rule::await_kw
            | Rule::break_kw
            | Rule::case_kw
            | Rule::catch_kw
            | Rule::class_kw
            | Rule::const_kw
            | Rule::continue_kw
            | Rule::covariant_kw
            | Rule::default_kw
            | Rule::deferred_kw
            | Rule::do_kw
            | Rule::dynamic_kw
            | Rule::else_kw
            | Rule::enum_kw
            | Rule::export_kw
            | Rule::extends_kw
            | Rule::extension_kw
            | Rule::external_kw
            | Rule::factory_kw
            | Rule::false_kw
            | Rule::final_kw
            | Rule::finally_kw
            | Rule::for_kw
            | Rule::function_kw
            | Rule::hide_kw
            | Rule::if_kw
            | Rule::implements_kw
            | Rule::import_kw
            | Rule::in_kw
            | Rule::interface_kw
            | Rule::is_kw
            | Rule::late_kw
            | Rule::library_kw
            | Rule::mixin_kw
            | Rule::native_kw
            | Rule::new_kw
            | Rule::null_kw
            | Rule::on_kw
            | Rule::operator_kw
            | Rule::override_kw
            | Rule::part_kw
            | Rule::required_kw
            | Rule::rethrow_kw
            | Rule::return_kw
            | Rule::show_kw
            | Rule::static_kw
            | Rule::super_kw
            | Rule::switch_kw
            | Rule::sync_kw
            | Rule::this_kw
            | Rule::throw_kw
            | Rule::true_kw
            | Rule::try_kw
            | Rule::typedef_kw
            | Rule::var_keyword
            | Rule::void_kw
            | Rule::when_kw
            | Rule::while_kw
            | Rule::with_kw
            | Rule::yield_kw
    )
}

fn extract_type_name(pair: &Pair<Rule>) -> String {
    // Extract the base type name from a type_annotation, stripping generics and nullable
    let s = pair.as_str().trim();
    let without_nullable = s.trim_end_matches('?').trim();
    common_generics::generic_base_name(without_nullable).to_string()
}

fn extract_type_name_from_clause(pair: &Pair<Rule>) -> Option<String> {
    for p in pair.clone().into_inner() {
        if p.as_rule() == Rule::type_annotation {
            return Some(extract_type_name(&p));
        }
    }
    None
}

fn extract_type_from_inner(pair: Pair<Rule>) -> String {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::type_annotation {
            return p.as_str().trim().to_string();
        }
    }
    "dynamic".to_string()
}

fn unquote_string_literal(pair: &Pair<Rule>) -> String {
    let s = pair.as_str();
    // Handle raw strings
    if s.starts_with("r'") || s.starts_with("r\"") {
        return s[2..s.len() - 1].to_string();
    }
    // Handle triple-quoted strings
    if s.starts_with("'''") || s.starts_with("\"\"\"") {
        return unescape_string_chars(&s[3..s.len() - 3]);
    }
    // Handle single/double quoted
    if s.len() >= 2 {
        return unescape_string_chars(&s[1..s.len() - 1]);
    }
    s.to_string()
}

fn unescape_string_chars(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let mut chars = s.chars();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.next() {
                Some('n') => result.push('\n'),
                Some('t') => result.push('\t'),
                Some('r') => result.push('\r'),
                Some('\\') => result.push('\\'),
                Some('\'') => result.push('\''),
                Some('"') => result.push('"'),
                Some('$') => result.push('$'),
                Some('0') => result.push('\0'),
                Some('x') => {
                    let hi = chars.next();
                    let lo = chars.next();
                    if let (Some(hi), Some(lo)) = (hi, lo) {
                        let hex = format!("{}{}", hi, lo);
                        if let Ok(value) = u32::from_str_radix(&hex, 16) {
                            if let Some(ch) = char::from_u32(value) {
                                result.push(ch);
                            }
                        }
                    }
                }
                Some('u') => {
                    let mut hex = String::new();
                    for _ in 0..4 {
                        if let Some(ch) = chars.next() {
                            hex.push(ch);
                        }
                    }
                    if let Ok(value) = u32::from_str_radix(&hex, 16) {
                        if let Some(ch) = char::from_u32(value) {
                            result.push(ch);
                        }
                    }
                }
                Some(other) => {
                    result.push('\\');
                    result.push(other);
                }
                None => result.push('\\'),
            }
        } else {
            result.push(c);
        }
    }
    result
}

fn compound_to_binop(op: CompoundOp) -> BinOp {
    match op {
        CompoundOp::Add => BinOp::Add,
        CompoundOp::Sub => BinOp::Sub,
        CompoundOp::Mul => BinOp::Mul,
        CompoundOp::Div => BinOp::Div,
        CompoundOp::IDiv => BinOp::IDiv,
        CompoundOp::Mod => BinOp::Mod,
        CompoundOp::Pow => BinOp::Pow,
        CompoundOp::BitAnd => BinOp::BitAnd,
        CompoundOp::BitOr => BinOp::BitOr,
        CompoundOp::BitXor => BinOp::BitXor,
        CompoundOp::Shl => BinOp::Shl,
        CompoundOp::Shr => BinOp::Shr,
        CompoundOp::UShr => BinOp::UShr,
        CompoundOp::And => BinOp::And,
        CompoundOp::Or => BinOp::Or,
        CompoundOp::NullCoalesce => BinOp::NullCoalesce,
        CompoundOp::Concat => BinOp::Concat,
    }
}
