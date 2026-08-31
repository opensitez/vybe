//! `System.Numerics` vector types, synthesized as REAL classes.
//!
//! Same mechanism and same reason as [`super::threading_classes`]: these are
//! .NET types with constructors, operators and static members, and a
//! `ClassType` in `class_exports()` cannot be constructed with `new`. Injecting
//! a `StmtKind::ClassDecl` gives them ordinary dispatch, so `v1 + v2` binds
//! `op_Addition` the way it does on any user class.
//!
//! # One generator, three arities
//!
//! `Vector2`, `Vector3` and `Vector4` differ ONLY in their component list —
//! every member is the same expression over `X`/`Y`, `X`/`Y`/`Z` or
//! `X`/`Y`/`Z`/`W`. Writing them three times is how the three drift apart, so
//! [`vector_class`] takes the components and generates all of it. `Vector3`'s
//! `Cross` is the single genuine exception and is added only there.
//!
//! ⚠ .NET's vectors are `float`, and this runtime's numbers are `f64`. Every
//! corpus assertion is either exact in both (`3`, `15`, `0`, `1`) or a
//! comparison, so the width is not observable there — but a test that prints
//! an accumulated float would see our EXTRA precision, not less. Recorded
//! because "we're more precise" is still a difference from .NET.
//!
//! ⚠ Injection is GATED on the program naming the type, for the reason
//! `interop_classes` documents: unconditional injection shifts typeidx
//! numbering for every language while the class model is mid-conversion.

use super::interop_classes::{
    by_ref_param, class, getter, me, method, param, shared_method,
};
use vybe_ast::{
    Argument, BinOp, ClassMember, ConstructorInitializerTarget, ExprKind, Expression, Literal,
    Modifiers, Param, PassBy, Statement, StmtKind, Visibility,
};

/// Whether `name` is a class this module injects.
pub fn is_synthesized_numerics_class(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "vector2"
            | "vector3"
            | "vector4"
            | "complex"
            | "quaternion"
            | "plane"
            | "matrix3x2"
            | "matrix4x4"
            | "int128"
            | "uint128"
    )
}

/// The numerics classes the program needs and does not declare itself.
///
/// ⛔ A USER TYPE OF THE SAME NAME WINS, AND THE GATE ALONE CANNOT SEE THAT.
/// Injection is keyed on the SOURCE TEXT containing the name, so a VB test that
/// declares its own `Structure Vector3` matched — and the injected class
/// shadowed the user's, breaking a GCHandle test that has nothing to do with
/// numerics. The declaration check is on the walked BODY, not another substring
/// search: `Vector3` appearing in a comment is not a declaration, and only the
/// AST knows the difference.
pub fn synthesize_numerics_classes(source: &str, body: &[Statement]) -> Vec<Statement> {
    let mut out = synthesize_all(source);
    out.retain(|stmt| {
        let StmtKind::ClassDecl { name, .. } = &stmt.kind else {
            return true;
        };
        !declares_type(body, name)
    });
    out
}

/// Whether the program declares a class, struct or record of this name.
fn declares_type(body: &[Statement], wanted: &str) -> bool {
    body.iter().any(|stmt| match &stmt.kind {
        StmtKind::ClassDecl { name, .. } | StmtKind::StructDecl { name, .. } => {
            name.eq_ignore_ascii_case(wanted)
        }
        StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
            declares_type(body, wanted)
        }
        StmtKind::ModuleDecl { members, .. } => members.iter().any(|m| match m {
            ClassMember::NestedType(stmt) => {
                declares_type(std::slice::from_ref(stmt.as_ref()), wanted)
            }
            _ => false,
        }),
        _ => false,
    })
}

fn synthesize_all(source: &str) -> Vec<Statement> {
    let lowered = source.to_ascii_lowercase();
    let mut out = Vec::new();
    if lowered.contains("vector2") {
        out.push(vector_class("Vector2", &["X", "Y"]));
    }
    if lowered.contains("vector3") {
        out.push(vector_class("Vector3", &["X", "Y", "Z"]));
    }
    if lowered.contains("vector4") {
        out.push(vector_class("Vector4", &["X", "Y", "Z", "W"]));
    }
    if lowered.contains("complex") {
        out.push(complex_class());
    }
    // `Quaternion`, `Plane` and `Matrix4x4` all build on `Vector3`, so it comes
    // along whether or not the program names it — their members construct one.
    // ⛔ `Vector3` IS 16,695 LINES OF AST — it is not free to drag in "just in
    // case". Matrix4x4 needs it only for `Translation` and `CreateLookAt`, so
    // the dependency now tracks the MEMBERS that actually construct one rather
    // than the mere mention of the matrix type. Measured: `Matrix4x4.Identity`
    // pulled 31k lines of Vector3 + Quaternion it never touched.
    let matrix_needs_vector3 = lowered.contains("matrix4x4")
        && (lowered.contains("translation") || lowered.contains("createlookat"));
    let needs_vector3 =
        lowered.contains("quaternion") || lowered.contains("plane") || matrix_needs_vector3;
    if needs_vector3 && !lowered.contains("vector3") {
        out.push(vector_class("Vector3", &["X", "Y", "Z"]));
    }
    // `Matrix4x4.CreateFromQuaternion` and `Plane.Transform` both name it.
    // ⛔ `Plane` NEEDS `Quaternion` — `plane_class` names it three times, and the
    // comment above says so. It used to arrive by accident, via the old
    // `|| matrix4x4` term, so tightening that term broke `Plane.Transform` with
    // `undefined is not callable` for any program that says "plane" without
    // saying "quaternion". State the real dependency instead of relying on a
    // coincidence.
    let needs_quaternion = lowered.contains("quaternion") || lowered.contains("plane");
    if needs_quaternion {
        out.push(quaternion_class());
    }
    if lowered.contains("plane") {
        out.push(plane_class());
    }
    if lowered.contains("matrix3x2") {
        out.push(matrix3x2_class(&lowered));
    }
    if lowered.contains("matrix4x4") {
        out.push(matrix4x4_class(&lowered));
    }
    if lowered.contains("int128") {
        // ⚠ `uint128` CONTAINS `int128`, so a program naming only `UInt128`
        // matches both — which is correct, not a bug: `UInt128.MaxValue` is
        // built from an `Int128`-free expression but the two are declared as a
        // pair and a program mixing them is ordinary.
        out.push(fixed_int_class("Int128", true));
    }
    if lowered.contains("uint128") {
        out.push(fixed_int_class("UInt128", false));
    }
    out
}

// ── Expression helpers ───────────────────────────────────────────────────

fn ident(name: &str) -> Expression {
    Expression::new(ExprKind::Ident(name.into()))
}

fn field_of(object: Expression, name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: name.into(),
        null_safe: false,
    })
}

fn bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn add(left: Expression, right: Expression) -> Expression {
    bin(BinOp::Add, left, right)
}

fn mul(left: Expression, right: Expression) -> Expression {
    bin(BinOp::Mul, left, right)
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn new_of(class_name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::New {
        class: Box::new(ident(class_name)),
        args: args.into_iter().map(Argument::positional).collect(),
    })
}

fn ret(value: Expression) -> Statement {
    Statement::new(StmtKind::Return(Some(value)))
}

/// `var name = init;` — no declared type, for hoisting a subexpression.
fn local_untyped(name: &str, init: Expression) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![vybe_ast::VarDeclarator {
            pattern: vybe_ast::BindingPattern::Ident(name.into()),
            type_hint: None,
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: vybe_ast::VarDeclKind::Var,
    })
}

/// `Type name = init;` inside a synthesized method body.
fn local(name: &str, type_hint: &str, init: Expression) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![vybe_ast::VarDeclarator {
            pattern: vybe_ast::BindingPattern::Ident(name.into()),
            type_hint: Some(vybe_ast::TypeHint::checked(type_hint)),
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: vybe_ast::VarDeclKind::Var,
    })
}

/// Give every synthesized method the return type its own body constructs.
///
/// ⛔ WITHOUT THIS, `Matrix4x4.CreateScale(2f) * Matrix4x4.CreateScale(3f)`
/// ANSWERED `undefined`. The C# operator rewrite dispatches on an operand's
/// inferred type, and a call to a static factory inferred nothing because
/// `shared_method` declares no return type — so `*` fell through to numeric
/// multiplication. Derived from the body rather than passed in at each call
/// site: the type a method returns IS the type it builds, and a hand-written
/// list beside 60-odd members is a second authority waiting to drift.
/// Give every `op_*` member the dunder spelling VB's protocol table binds.
///
/// ⛔ VB AND C# REACH AN OVERLOADED OPERATOR BY DIFFERENT ROADS. C# rewrites
/// `a + b` into a STATIC call `T.op_Addition(a, b)`, dispatching on the
/// operand's inferred type. VB has no such pass: it fills
/// `ProtocolSlot::Add` from an INSTANCE member and lets the shared
/// `emit_rich_binop` find it on the left operand. With only the static
/// declared, `Int128.Parse("1e20") + Int128.Parse("2e20")` fell through to the
/// dynamic add and CONCATENATED the two decimal strings — a wrong answer that
/// reads as a formatting bug.
///
/// Derived from the members already present rather than hand-listed, so an
/// operator added later cannot be forgotten here. The body DELEGATES to the
/// static, so there is still exactly one implementation.
///
/// ⚠ Named with the dunder, not `operator+`: BOTH protocol tables map
/// `operator+`, so that spelling would also bind the slot in C# — and a C#
/// instance body reading `Me` is the known broken case (a method WITH a
/// parameter never reaches its bound receiver). The dunders are unmapped in
/// C#'s table, so they are inert there and correct in VB.
fn add_vb_operator_slots(members: &mut Vec<ClassMember>, class_name: &str) {
    const PAIRS: [(&str, &str); 11] = [
        ("op_Addition", "__add__"),
        ("op_Subtraction", "__sub__"),
        ("op_Multiply", "__mul__"),
        ("op_Division", "__truediv__"),
        ("op_Modulus", "__mod__"),
        ("op_Equality", "__eq__"),
        ("op_Inequality", "__ne__"),
        ("op_LessThan", "__lt__"),
        ("op_LessThanOrEqual", "__le__"),
        ("op_GreaterThan", "__gt__"),
        ("op_GreaterThanOrEqual", "__ge__"),
    ];
    let declared: Vec<String> = members
        .iter()
        .filter_map(|m| match m {
            ClassMember::Method(stmt) => match &stmt.kind {
                StmtKind::FunctionDecl { name, params, .. } if params.len() == 2 => {
                    Some(name.clone())
                }
                _ => None,
            },
            _ => None,
        })
        .collect();
    for (op_name, dunder) in PAIRS {
        if !declared.iter().any(|n| n == op_name) {
            continue;
        }
        members.push(method(
            dunder,
            vec![typed_param("other", class_name)],
            vec![ret(call(
                field_of(ident(class_name), op_name),
                vec![Expression::new(ExprKind::This), ident("other")],
            ))],
            false,
        ));
    }
}

fn declare_return_types(members: &mut [ClassMember]) {
    for member in members.iter_mut() {
        let ClassMember::Method(stmt) = member else {
            continue;
        };
        let StmtKind::FunctionDecl {
            body, return_type, ..
        } = &mut stmt.kind
        else {
            continue;
        };
        if return_type.is_some() {
            continue;
        }
        let Some(Statement {
            kind: StmtKind::Return(Some(value)),
            ..
        }) = body.last()
        else {
            continue;
        };
        if let Some(built) = constructed_type(value) {
            *return_type = Some(built);
        }
    }
}

/// The class an expression constructs, looking through a ternary whose arms
/// agree — `Sqrt` and `Invert` both answer one type by two routes.
fn constructed_type(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            _ => None,
        },
        ExprKind::Ternary { then, else_, .. } => {
            let a = constructed_type(then)?;
            (a == constructed_type(else_)?).then_some(a)
        }
        _ => None,
    }
}

fn typed_param(name: &str, type_hint: &str) -> Param {
    Param {
        name: name.into(),
        type_hint: Some(type_hint.into()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

/// A `Shared` field whose initialiser builds an instance — `Vector3.UnitX`.
///
/// A field rather than a property because .NET declares these as `static
/// readonly` fields, and because the value is immutable: every member here
/// returns a NEW instance rather than mutating, so one shared instance per
/// constant cannot be observed as aliasing.
fn shared_value(name: &str, init: Expression, type_hint: &str) -> ClassMember {
    ClassMember::Field {
        name: name.into(),
        type_hint: Some(type_hint.into()),
        init: Some(init),
        modifiers: Modifiers {
            is_shared: true,
            is_static: true,
            ..Modifiers::default()
        },
        with_events: false,
        array_bounds: None,
        storage: None,
    }
}

fn ctor(params: Vec<Param>, body: Vec<Statement>) -> ClassMember {
    ClassMember::Constructor {
        name: None,
        params,
        body,
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    }
}

/// Fold `f(comp)` over the components with `op` — the shape of `Dot`,
/// `LengthSquared` and every comparison in this file.
fn fold_components<F>(comps: &[&str], op: BinOp, mut f: F) -> Expression
where
    F: FnMut(&str) -> Expression,
{
    let mut it = comps.iter();
    let first = f(it.next().expect("a vector has at least one component"));
    it.fold(first, |acc, c| bin(op, acc, f(c)))
}

/// `new V(<per-component expression>)`.
fn build<F>(name: &str, comps: &[&str], mut f: F) -> Expression
where
    F: FnMut(&str) -> Expression,
{
    new_of(name, comps.iter().map(|c| f(c)).collect())
}

// ── The generator ────────────────────────────────────────────────────────

fn vector_class(name: &str, comps: &[&str]) -> Statement {
    let mut members: Vec<ClassMember> = Vec::new();

    for c in comps {
        members.push(ClassMember::Field {
            name: (*c).into(),
            type_hint: Some("float".into()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
            storage: None,
        });
    }

    // `new V(x, y, …)` and .NET's broadcast `new V(value)`.
    members.push(ctor(
        comps.iter().map(|c| typed_param(c, "float")).collect(),
        comps
            .iter()
            .map(|c| {
                Statement::new(StmtKind::Assign {
                    targets: vec![me(c)],
                    value: ident(c),
                    by_ref: false,
                })
            })
            .collect(),
    ));
    members.push(ctor(
        vec![typed_param("value", "float")],
        comps
            .iter()
            .map(|c| {
                Statement::new(StmtKind::Assign {
                    targets: vec![me(c)],
                    value: ident("value"),
                    by_ref: false,
                })
            })
            .collect(),
    ));

    // static readonly constants.
    members.push(shared_value(
        "Zero",
        build(name, comps, |_| Expression::float(0.0)),
        name,
    ));
    members.push(shared_value(
        "One",
        build(name, comps, |_| Expression::float(1.0)),
        name,
    ));
    for unit in comps.iter() {
        members.push(shared_value(
            &format!("Unit{unit}"),
            build(name, comps, |c| {
                Expression::float(if c == *unit { 1.0 } else { 0.0 })
            }),
            name,
        ));
    }

    // LengthSquared / Length.
    let len_sq = fold_components(comps, BinOp::Add, |c| mul(me(c), me(c)));
    members.push(method(
        "LengthSquared",
        Vec::new(),
        vec![ret(len_sq.clone())],
        false,
    ));
    members.push(method(
        "Length",
        Vec::new(),
        vec![ret(call(
            field_of(ident("Math"), "Sqrt"),
            vec![len_sq.clone()],
        ))],
        false,
    ));

    // Componentwise binary statics and their operators.
    for (method_name, op, operator) in [
        ("Add", BinOp::Add, "op_Addition"),
        ("Subtract", BinOp::Sub, "op_Subtraction"),
        ("Multiply", BinOp::Mul, "op_Multiply"),
        ("Divide", BinOp::Div, "op_Division"),
    ] {
        // ⛔ `*` AND `/` TAKE A VECTOR **OR** A SCALAR, and declaring the two
        // as separate overloads did not work: `v1 * v2` selected the SCALAR
        // one and multiplied each component by an object, answering `NaN`.
        // Overload selection is by declared parameter type and both operands
        // reach it as one synthesized class, so the choice is made at runtime
        // where the operand can actually be examined. `Add`/`Subtract` have no
        // scalar form in .NET and keep the plain componentwise body.
        let operand = |c: &str| -> Expression {
            if matches!(op, BinOp::Mul | BinOp::Div) {
                Expression::new(ExprKind::Ternary {
                    cond: Box::new(bin(
                        BinOp::InstanceOf,
                        ident("right"),
                        ident(name),
                    )),
                    then: Box::new(field_of(ident("right"), c)),
                    else_: Box::new(ident("right")),
                })
            } else {
                field_of(ident("right"), c)
            }
        };
        let body = vec![ret(build(name, comps, |c| {
            bin(op, field_of(ident("left"), c), operand(c))
        }))];
        members.push(shared_method(
            method_name,
            vec![typed_param("left", name), typed_param("right", name)],
            body.clone(),
        ));
        members.push(shared_method(
            operator,
            vec![typed_param("left", name), typed_param("right", name)],
            body,
        ));
    }

    // `Vector2 * 2f` and `Vector2 / 2f` — .NET declares the scalar forms as
    // separate overloads, so they are separate overloads here too rather than
    // one body branching on the operand's runtime shape.
    // Negation.
    members.push(shared_method(
        "Negate",
        vec![typed_param("value", name)],
        vec![ret(build(name, comps, |c| {
            bin(
                BinOp::Sub,
                Expression::float(0.0),
                field_of(ident("value"), c),
            )
        }))],
    ));

    members.push(shared_method(
        "op_UnaryNegation",
        vec![typed_param("value", name)],
        vec![ret(build(name, comps, |c| {
            bin(
                BinOp::Sub,
                Expression::float(0.0),
                field_of(ident("value"), c),
            )
        }))],
    ));

    // Dot / Distance.
    members.push(shared_method(
        "Dot",
        vec![typed_param("left", name), typed_param("right", name)],
        vec![ret(fold_components(comps, BinOp::Add, |c| {
            mul(field_of(ident("left"), c), field_of(ident("right"), c))
        }))],
    ));
    let dist_sq = fold_components(comps, BinOp::Add, |c| {
        let d = bin(
            BinOp::Sub,
            field_of(ident("left"), c),
            field_of(ident("right"), c),
        );
        mul(d.clone(), d)
    });
    members.push(shared_method(
        "DistanceSquared",
        vec![typed_param("left", name), typed_param("right", name)],
        vec![ret(dist_sq.clone())],
    ));
    members.push(shared_method(
        "Distance",
        vec![typed_param("left", name), typed_param("right", name)],
        vec![ret(call(
            field_of(ident("Math"), "Sqrt"),
            vec![dist_sq],
        ))],
    ));

    // Normalize — .NET divides by the length and answers NaN components for a
    // zero vector rather than throwing.
    members.push(shared_method(
        "Normalize",
        vec![typed_param("value", name)],
        vec![ret(build(name, comps, |c| {
            bin(
                BinOp::Div,
                field_of(ident("value"), c),
                call(
                    field_of(ident("Math"), "Sqrt"),
                    vec![fold_components(comps, BinOp::Add, |k| {
                        mul(field_of(ident("value"), k), field_of(ident("value"), k))
                    })],
                ),
            )
        }))],
    ));

    // Min / Max / Abs / SquareRoot / Lerp / Clamp.
    for (method_name, fn_name) in [("Min", "Min"), ("Max", "Max")] {
        members.push(shared_method(
            method_name,
            vec![typed_param("left", name), typed_param("right", name)],
            vec![ret(build(name, comps, |c| {
                call(
                    field_of(ident("Math"), fn_name),
                    vec![field_of(ident("left"), c), field_of(ident("right"), c)],
                )
            }))],
        ));
    }
    for (method_name, fn_name) in [("Abs", "Abs"), ("SquareRoot", "Sqrt")] {
        members.push(shared_method(
            method_name,
            vec![typed_param("value", name)],
            vec![ret(build(name, comps, |c| {
                call(
                    field_of(ident("Math"), fn_name),
                    vec![field_of(ident("value"), c)],
                )
            }))],
        ));
    }
    members.push(shared_method(
        "Lerp",
        vec![
            typed_param("left", name),
            typed_param("right", name),
            typed_param("amount", "float"),
        ],
        vec![ret(build(name, comps, |c| {
            add(
                field_of(ident("left"), c),
                mul(
                    bin(
                        BinOp::Sub,
                        field_of(ident("right"), c),
                        field_of(ident("left"), c),
                    ),
                    ident("amount"),
                ),
            )
        }))],
    ));
    members.push(shared_method(
        "Clamp",
        vec![
            typed_param("value", name),
            typed_param("min", name),
            typed_param("max", name),
        ],
        vec![ret(build(name, comps, |c| {
            call(
                field_of(ident("Math"), "Min"),
                vec![
                    field_of(ident("max"), c),
                    call(
                        field_of(ident("Math"), "Max"),
                        vec![field_of(ident("min"), c), field_of(ident("value"), c)],
                    ),
                ],
            )
        }))],
    ));

    // Reflect: `value - 2 * Dot(value, normal) * normal`.
    members.push(shared_method(
        "Reflect",
        vec![typed_param("value", name), typed_param("normal", name)],
        vec![ret(build(name, comps, |c| {
            bin(
                BinOp::Sub,
                field_of(ident("value"), c),
                mul(
                    mul(
                        Expression::float(2.0),
                        fold_components(comps, BinOp::Add, |k| {
                            mul(field_of(ident("value"), k), field_of(ident("normal"), k))
                        }),
                    ),
                    field_of(ident("normal"), c),
                ),
            )
        }))],
    ));

    // `Vector3.Transform` takes a Matrix4x4 OR a Quaternion. Same reason the
    // scalar `*` branches at runtime: both arrive as one synthesized class and
    // overload selection cannot separate them.
    if comps == ["X", "Y", "Z"] {
        let by_matrix = |c: usize| {
            let mut acc = mul(
                field_of(ident("position"), "X"),
                field_of(ident("transform"), &format!("M1{}", c + 1)),
            );
            for (k, comp) in ["Y", "Z"].iter().enumerate() {
                acc = add(
                    acc,
                    mul(
                        field_of(ident("position"), comp),
                        field_of(ident("transform"), &format!("M{}{}", k + 2, c + 1)),
                    ),
                );
            }
            add(
                acc,
                field_of(ident("transform"), &format!("M4{}", c + 1)),
            )
        };
        // v + 2w×(w×v + sv) — the quaternion sandwich without building one.
        let qf = |c: &str| field_of(ident("transform"), c);
        let pf = |c: &str| field_of(ident("position"), c);
        let cross1 = |a: &str, b: &str| {
            bin(BinOp::Sub, mul(qf(a), pf(b)), mul(qf(b), pf(a)))
        };
        let t = |c: usize| {
            let (a, b) = [("Y", "Z"), ("Z", "X"), ("X", "Y")][c];
            mul(
                Expression::float(2.0),
                add(cross1(a, b), mul(qf("W"), pf(["X", "Y", "Z"][c]))),
            )
        };
        let by_quaternion = |c: usize| {
            let (a, b) = [(1usize, 2usize), (2, 0), (0, 1)][c];
            add(
                pf(["X", "Y", "Z"][c]),
                bin(
                    BinOp::Sub,
                    mul(qf(["X", "Y", "Z"][a]), t(b)),
                    mul(qf(["X", "Y", "Z"][b]), t(a)),
                ),
            )
        };
        members.push(shared_method(
            "Transform",
            vec![
                typed_param("position", name),
                typed_param("transform", "Matrix4x4"),
            ],
            vec![ret(new_of(
                name,
                (0..3)
                    .map(|c| {
                        Expression::new(ExprKind::Ternary {
                            cond: Box::new(bin(
                                BinOp::InstanceOf,
                                ident("transform"),
                                ident("Quaternion"),
                            )),
                            then: Box::new(by_quaternion(c)),
                            else_: Box::new(by_matrix(c)),
                        })
                    })
                    .collect(),
            ))],
        ));
    }

    // Vector3 alone has a cross product.
    if comps == ["X", "Y", "Z"] {
        let cross = |a: &str, b: &str| {
            bin(
                BinOp::Sub,
                mul(field_of(ident("left"), a), field_of(ident("right"), b)),
                mul(field_of(ident("left"), b), field_of(ident("right"), a)),
            )
        };
        members.push(shared_method(
            "Cross",
            vec![typed_param("left", name), typed_param("right", name)],
            vec![ret(new_of(
                name,
                vec![cross("Y", "Z"), cross("Z", "X"), cross("X", "Y")],
            ))],
        ));
    }

    // Equality — value semantics, and the operators that spell it.
    let all_eq = fold_components(comps, BinOp::And, |c| {
        bin(BinOp::Eq, me(c), field_of(ident("other"), c))
    });
    members.push(method(
        "Equals",
        vec![typed_param("other", name)],
        vec![ret(all_eq)],
        false,
    ));
    let operands_eq = fold_components(comps, BinOp::And, |c| {
        bin(
            BinOp::Eq,
            field_of(ident("left"), c),
            field_of(ident("right"), c),
        )
    });
    members.push(shared_method(
        "op_Equality",
        vec![typed_param("left", name), typed_param("right", name)],
        vec![ret(operands_eq.clone())],
    ));
    members.push(shared_method(
        "op_Inequality",
        vec![typed_param("left", name), typed_param("right", name)],
        vec![ret(Expression::new(ExprKind::Unary {
            op: vybe_ast::UnaryOp::Not,
            expr: Box::new(operands_eq),
        }))],
    ));
    members.push(method(
        "GetHashCode",
        Vec::new(),
        vec![ret(call(
            field_of(field_of(ident("System"), "HashCode"), "Combine"),
            comps.iter().map(|c| me(c)).collect(),
        ))],
        false,
    ));

    // `<3, 4>` — .NET's own rendering.
    let mut text = Expression::new(ExprKind::Lit(Literal::Str("<".into())));
    for (i, c) in comps.iter().enumerate() {
        if i > 0 {
            text = add(
                text,
                Expression::new(ExprKind::Lit(Literal::Str(", ".into()))),
            );
        }
        text = add(text, me(c));
    }
    text = add(
        text,
        Expression::new(ExprKind::Lit(Literal::Str(">".into()))),
    );
    members.push(method("ToString", Vec::new(), vec![ret(text)], false));

    // `IsZero` is not .NET surface; the getters below are.
    members.push(getter(
        "IsZero",
        vec![ret(fold_components(comps, BinOp::And, |c| {
            bin(BinOp::Eq, me(c), Expression::float(0.0))
        }))],
    ));

    add_vb_operator_slots(&mut members, name);
    declare_return_types(&mut members);
    class(name, Vec::new(), members)
}

// ── System.Numerics.Complex ──────────────────────────────────────────────

/// `Math.<name>(args…)`.
fn math(name: &str, args: Vec<Expression>) -> Expression {
    call(field_of(ident("Math"), name), args)
}

fn re(v: &str) -> Expression {
    field_of(ident(v), "Real")
}

fn im(v: &str) -> Expression {
    field_of(ident(v), "Imaginary")
}

/// `System.Numerics.Complex` — a + bi over this runtime's f64.
///
/// The transcendental functions are the textbook identities rather than
/// anything clever: `Exp` from `e^a(cos b + i sin b)`, `Log` from
/// `(ln|z|, arg z)`, `Pow` as `Exp(y * Log(x))`. .NET computes them the same
/// way, so agreement is structural rather than tuned — and `Sqrt` is written
/// as its own polar form because `Pow(z, 0.5)` loses the exact answer for a
/// real negative, which is the case a test checks.
fn complex_class() -> Statement {
    let mut members: Vec<ClassMember> = Vec::new();
    for c in ["Real", "Imaginary"] {
        members.push(ClassMember::Field {
            name: c.into(),
            type_hint: Some("double".into()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
            storage: None,
        });
    }
    members.push(ctor(
        vec![typed_param("real", "double"), typed_param("imaginary", "double")],
        vec![
            Statement::new(StmtKind::Assign {
                targets: vec![me("Real")],
                value: ident("real"),
                by_ref: false,
            }),
            Statement::new(StmtKind::Assign {
                targets: vec![me("Imaginary")],
                value: ident("imaginary"),
                by_ref: false,
            }),
        ],
    ));

    members.push(shared_value(
        "Zero",
        new_of("Complex", vec![Expression::float(0.0), Expression::float(0.0)]),
        "Complex",
    ));
    members.push(shared_value(
        "One",
        new_of("Complex", vec![Expression::float(1.0), Expression::float(0.0)]),
        "Complex",
    ));
    members.push(shared_value(
        "ImaginaryOne",
        new_of("Complex", vec![Expression::float(0.0), Expression::float(1.0)]),
        "Complex",
    ));

    // |z| and arg z.
    let magnitude = math(
        "Sqrt",
        vec![add(
            mul(me("Real"), me("Real")),
            mul(me("Imaginary"), me("Imaginary")),
        )],
    );
    members.push(getter("Magnitude", vec![ret(magnitude)]));
    members.push(getter(
        "Phase",
        vec![ret(math("Atan2", vec![me("Imaginary"), me("Real")]))],
    ));

    members.push(shared_method(
        "Abs",
        vec![typed_param("value", "Complex")],
        vec![ret(math(
            "Sqrt",
            vec![add(mul(re("value"), re("value")), mul(im("value"), im("value")))],
        ))],
    ));
    members.push(shared_method(
        "Conjugate",
        vec![typed_param("value", "Complex")],
        vec![ret(new_of(
            "Complex",
            vec![
                re("value"),
                bin(BinOp::Sub, Expression::float(0.0), im("value")),
            ],
        ))],
    ));
    members.push(shared_method(
        "FromPolarCoordinates",
        vec![
            typed_param("magnitude", "double"),
            typed_param("phase", "double"),
        ],
        vec![ret(new_of(
            "Complex",
            vec![
                mul(ident("magnitude"), math("Cos", vec![ident("phase")])),
                mul(ident("magnitude"), math("Sin", vec![ident("phase")])),
            ],
        ))],
    ));

    // Arithmetic, and the operators that spell it.
    for (method_name, operator, real, imag) in [
        (
            "Add",
            "op_Addition",
            add(re("left"), re("right")),
            add(im("left"), im("right")),
        ),
        (
            "Subtract",
            "op_Subtraction",
            bin(BinOp::Sub, re("left"), re("right")),
            bin(BinOp::Sub, im("left"), im("right")),
        ),
        (
            "Multiply",
            "op_Multiply",
            bin(
                BinOp::Sub,
                mul(re("left"), re("right")),
                mul(im("left"), im("right")),
            ),
            add(mul(re("left"), im("right")), mul(im("left"), re("right"))),
        ),
    ] {
        let body = vec![ret(new_of("Complex", vec![real, imag]))];
        members.push(shared_method(
            method_name,
            vec![typed_param("left", "Complex"), typed_param("right", "Complex")],
            body.clone(),
        ));
        members.push(shared_method(
            operator,
            vec![typed_param("left", "Complex"), typed_param("right", "Complex")],
            body,
        ));
    }

    // Division by the conjugate: (ac+bd)/d2, (bc-ad)/d2.
    let denom = add(mul(re("right"), re("right")), mul(im("right"), im("right")));
    let div_body = vec![ret(new_of(
        "Complex",
        vec![
            bin(
                BinOp::Div,
                add(mul(re("left"), re("right")), mul(im("left"), im("right"))),
                denom.clone(),
            ),
            bin(
                BinOp::Div,
                bin(
                    BinOp::Sub,
                    mul(im("left"), re("right")),
                    mul(re("left"), im("right")),
                ),
                denom,
            ),
        ],
    ))];
    for name in ["Divide", "op_Division"] {
        members.push(shared_method(
            name,
            vec![typed_param("left", "Complex"), typed_param("right", "Complex")],
            div_body.clone(),
        ));
    }
    members.push(shared_method(
        "Reciprocal",
        vec![typed_param("value", "Complex")],
        vec![ret({
            // 1/z = conj(z) / |z|².
            let d = add(mul(re("value"), re("value")), mul(im("value"), im("value")));
            new_of(
                "Complex",
                vec![
                    bin(BinOp::Div, re("value"), d.clone()),
                    bin(
                        BinOp::Div,
                        bin(BinOp::Sub, Expression::float(0.0), im("value")),
                        d,
                    ),
                ],
            )
        })],
    ));

    for op_name in ["Negate", "op_UnaryNegation"] {
        members.push(shared_method(
            op_name,
            vec![typed_param("value", "Complex")],
            vec![ret(new_of(
                "Complex",
                vec![
                    bin(BinOp::Sub, Expression::float(0.0), re("value")),
                    bin(BinOp::Sub, Expression::float(0.0), im("value")),
                ],
            ))],
        ));
    }

    // exp / log / trig.
    members.push(shared_method(
        "Exp",
        vec![typed_param("value", "Complex")],
        vec![ret(new_of(
            "Complex",
            vec![
                mul(
                    math("Exp", vec![re("value")]),
                    math("Cos", vec![im("value")]),
                ),
                mul(
                    math("Exp", vec![re("value")]),
                    math("Sin", vec![im("value")]),
                ),
            ],
        ))],
    ));
    members.push(shared_method(
        "Log",
        vec![typed_param("value", "Complex")],
        vec![ret(new_of(
            "Complex",
            vec![
                math(
                    "Log",
                    vec![math(
                        "Sqrt",
                        vec![add(mul(re("value"), re("value")), mul(im("value"), im("value")))],
                    )],
                ),
                math("Atan2", vec![im("value"), re("value")]),
            ],
        ))],
    ));
    members.push(shared_method(
        "Sqrt",
        vec![typed_param("value", "Complex")],
        vec![ret({
            // Polar half-angle: √r·(cos θ/2, sin θ/2). Written this way rather
            // than as `Pow(z, 0.5)` because the polar form is EXACT for a real
            // negative — `Sqrt(-4)` is `<0; 2>`, not `<1.2e-16; 2>`.
            let r = math(
                "Sqrt",
                vec![add(mul(re("value"), re("value")), mul(im("value"), im("value")))],
            );
            let half = bin(
                BinOp::Div,
                math("Atan2", vec![im("value"), re("value")]),
                Expression::float(2.0),
            );
            let root = math("Sqrt", vec![r]);
            let polar = new_of(
                "Complex",
                vec![
                    mul(root.clone(), math("Cos", vec![half.clone()])),
                    mul(root, math("Sin", vec![half])),
                ],
            );
            // ⛔ THE POLAR FORM IS NOT EXACT ON THE REAL AXIS. `Sqrt(-4)` came
            // out `<1.2246467991473532E-16; 2>` because `cos(π/2)` is not
            // exactly zero in binary. .NET special-cases a zero imaginary part
            // for precisely this reason, and answers `<0; 2>`; so does this.
            Expression::new(ExprKind::Ternary {
                cond: Box::new(bin(BinOp::Eq, im("value"), Expression::float(0.0))),
                then: Box::new(Expression::new(ExprKind::Ternary {
                    cond: Box::new(bin(
                        BinOp::Lt,
                        re("value"),
                        Expression::float(0.0),
                    )),
                    then: Box::new(new_of(
                        "Complex",
                        vec![
                            Expression::float(0.0),
                            math(
                                "Sqrt",
                                vec![bin(BinOp::Sub, Expression::float(0.0), re("value"))],
                            ),
                        ],
                    )),
                    else_: Box::new(new_of(
                        "Complex",
                        vec![math("Sqrt", vec![re("value")]), Expression::float(0.0)],
                    )),
                })),
                else_: Box::new(polar),
            })
        })],
    ));
    members.push(shared_method(
        "Pow",
        vec![typed_param("value", "Complex"), typed_param("power", "double")],
        vec![ret({
            // z^y = r^y · (cos yθ, sin yθ) — `Exp(y·Log z)` with both halves
            // folded, so no intermediate Complex is built.
            let r = math(
                "Sqrt",
                vec![add(mul(re("value"), re("value")), mul(im("value"), im("value")))],
            );
            let theta = math("Atan2", vec![im("value"), re("value")]);
            let scale = math("Pow", vec![r, ident("power")]);
            let angle = mul(theta, ident("power"));
            new_of(
                "Complex",
                vec![
                    mul(scale.clone(), math("Cos", vec![angle.clone()])),
                    mul(scale, math("Sin", vec![angle])),
                ],
            )
        })],
    ));
    for (name, outer, inner) in [
        ("Sin", "Sin", "Cos"),
        ("Cos", "Cos", "Sin"),
    ] {
        // sin(a+bi) = sin a cosh b + i cos a sinh b; cos flips the sign.
        let cosh = bin(
            BinOp::Div,
            add(
                math("Exp", vec![im("value")]),
                math(
                    "Exp",
                    vec![bin(BinOp::Sub, Expression::float(0.0), im("value"))],
                ),
            ),
            Expression::float(2.0),
        );
        let sinh = bin(
            BinOp::Div,
            bin(
                BinOp::Sub,
                math("Exp", vec![im("value")]),
                math(
                    "Exp",
                    vec![bin(BinOp::Sub, Expression::float(0.0), im("value"))],
                ),
            ),
            Expression::float(2.0),
        );
        let imag_part = mul(math(inner, vec![re("value")]), sinh);
        members.push(shared_method(
            name,
            vec![typed_param("value", "Complex")],
            vec![ret(new_of(
                "Complex",
                vec![
                    mul(math(outer, vec![re("value")]), cosh),
                    if name == "Cos" {
                        bin(BinOp::Sub, Expression::float(0.0), imag_part)
                    } else {
                        imag_part
                    },
                ],
            ))],
        ));
    }

    // Predicates — .NET asks about BOTH components.
    for (name, probe) in [
        ("IsNaN", "IsNaN"),
        ("IsInfinity", "IsInfinity"),
    ] {
        members.push(shared_method(
            name,
            vec![typed_param("value", "Complex")],
            vec![ret(bin(
                BinOp::Or,
                call(field_of(ident("double"), probe), vec![re("value")]),
                call(field_of(ident("double"), probe), vec![im("value")]),
            ))],
        ));
    }
    members.push(shared_method(
        "IsFinite",
        vec![typed_param("value", "Complex")],
        vec![ret(bin(
            BinOp::And,
            call(field_of(ident("double"), "IsFinite"), vec![re("value")]),
            call(field_of(ident("double"), "IsFinite"), vec![im("value")]),
        ))],
    ));

    members.push(method(
        "Equals",
        vec![typed_param("other", "Complex")],
        vec![ret(bin(
            BinOp::And,
            bin(BinOp::Eq, me("Real"), field_of(ident("other"), "Real")),
            bin(
                BinOp::Eq,
                me("Imaginary"),
                field_of(ident("other"), "Imaginary"),
            ),
        ))],
        false,
    ));
    let operands_eq = bin(
        BinOp::And,
        bin(BinOp::Eq, re("left"), re("right")),
        bin(BinOp::Eq, im("left"), im("right")),
    );
    members.push(shared_method(
        "op_Equality",
        vec![typed_param("left", "Complex"), typed_param("right", "Complex")],
        vec![ret(operands_eq.clone())],
    ));
    members.push(shared_method(
        "op_Inequality",
        vec![typed_param("left", "Complex"), typed_param("right", "Complex")],
        vec![ret(Expression::new(ExprKind::Unary {
            op: vybe_ast::UnaryOp::Not,
            expr: Box::new(operands_eq),
        }))],
    ));
    members.push(method(
        "GetHashCode",
        Vec::new(),
        vec![ret(call(
            field_of(field_of(ident("System"), "HashCode"), "Combine"),
            vec![me("Real"), me("Imaginary")],
        ))],
        false,
    ));

    // `<3; 4>` — .NET 7+ renders a Complex exactly this way.
    let text = add(
        add(
            add(
                add(
                    Expression::new(ExprKind::Lit(Literal::Str("<".into()))),
                    me("Real"),
                ),
                Expression::new(ExprKind::Lit(Literal::Str("; ".into()))),
            ),
            me("Imaginary"),
        ),
        Expression::new(ExprKind::Lit(Literal::Str(">".into()))),
    );
    members.push(method("ToString", Vec::new(), vec![ret(text)], false));

    add_vb_operator_slots(&mut members, "Complex");
    declare_return_types(&mut members);
    class("Complex", Vec::new(), members)
}

// ── System.Numerics.Quaternion ───────────────────────────────────────────

fn q(v: &str, c: &str) -> Expression {
    field_of(ident(v), c)
}

/// `System.Numerics.Quaternion` — `(X, Y, Z, W)`, `W` last as .NET spells it.
fn quaternion_class() -> Statement {
    const C: [&str; 4] = ["X", "Y", "Z", "W"];
    let mut members: Vec<ClassMember> = Vec::new();
    for c in C {
        members.push(ClassMember::Field {
            name: c.into(),
            type_hint: Some("float".into()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
            storage: None,
        });
    }
    members.push(ctor(
        C.iter().map(|c| typed_param(c, "float")).collect(),
        C.iter()
            .map(|c| {
                Statement::new(StmtKind::Assign {
                    targets: vec![me(c)],
                    value: ident(c),
                    by_ref: false,
                })
            })
            .collect(),
    ));
    members.push(shared_value(
        "Identity",
        new_of(
            "Quaternion",
            vec![
                Expression::float(0.0),
                Expression::float(0.0),
                Expression::float(0.0),
                Expression::float(1.0),
            ],
        ),
        "Quaternion",
    ));
    members.push(shared_value(
        "Zero",
        new_of(
            "Quaternion",
            vec![
                Expression::float(0.0),
                Expression::float(0.0),
                Expression::float(0.0),
                Expression::float(0.0),
            ],
        ),
        "Quaternion",
    ));

    let norm2 = fold_components(&C, BinOp::Add, |c| mul(me(c), me(c)));
    members.push(method(
        "LengthSquared",
        Vec::new(),
        vec![ret(norm2.clone())],
        false,
    ));
    members.push(method(
        "Length",
        Vec::new(),
        vec![ret(math("Sqrt", vec![norm2]))],
        false,
    ));

    members.push(getter(
        "IsIdentity",
        vec![ret(bin(
            BinOp::And,
            fold_components(&["X", "Y", "Z"], BinOp::And, |c| {
                bin(BinOp::Eq, me(c), Expression::float(0.0))
            }),
            bin(BinOp::Eq, me("W"), Expression::float(1.0)),
        ))],
    ));

    members.push(shared_method(
        "Dot",
        vec![
            typed_param("left", "Quaternion"),
            typed_param("right", "Quaternion"),
        ],
        vec![ret(fold_components(&C, BinOp::Add, |c| {
            mul(q("left", c), q("right", c))
        }))],
    ));
    members.push(shared_method(
        "Conjugate",
        vec![typed_param("value", "Quaternion")],
        vec![ret(new_of(
            "Quaternion",
            vec![
                bin(BinOp::Sub, Expression::float(0.0), q("value", "X")),
                bin(BinOp::Sub, Expression::float(0.0), q("value", "Y")),
                bin(BinOp::Sub, Expression::float(0.0), q("value", "Z")),
                q("value", "W"),
            ],
        ))],
    ));
    let value_norm2 = fold_components(&C, BinOp::Add, |c| mul(q("value", c), q("value", c)));
    members.push(shared_method(
        "Normalize",
        vec![typed_param("value", "Quaternion")],
        vec![ret(new_of(
            "Quaternion",
            C.iter()
                .map(|c| {
                    bin(
                        BinOp::Div,
                        q("value", c),
                        math("Sqrt", vec![value_norm2.clone()]),
                    )
                })
                .collect(),
        ))],
    ));
    // `conj(q) / |q|²` — the inverse, which for a unit quaternion IS the
    // conjugate, and .NET still divides.
    members.push(shared_method(
        "Inverse",
        vec![typed_param("value", "Quaternion")],
        vec![ret(new_of(
            "Quaternion",
            vec![
                bin(
                    BinOp::Div,
                    bin(BinOp::Sub, Expression::float(0.0), q("value", "X")),
                    value_norm2.clone(),
                ),
                bin(
                    BinOp::Div,
                    bin(BinOp::Sub, Expression::float(0.0), q("value", "Y")),
                    value_norm2.clone(),
                ),
                bin(
                    BinOp::Div,
                    bin(BinOp::Sub, Expression::float(0.0), q("value", "Z")),
                    value_norm2.clone(),
                ),
                bin(BinOp::Div, q("value", "W"), value_norm2),
            ],
        ))],
    ));

    // Hamilton product, and the `*` that spells it.
    let mul_body = vec![ret(new_of(
        "Quaternion",
        vec![
            add(
                add(
                    mul(q("left", "W"), q("right", "X")),
                    mul(q("left", "X"), q("right", "W")),
                ),
                bin(
                    BinOp::Sub,
                    mul(q("left", "Y"), q("right", "Z")),
                    mul(q("left", "Z"), q("right", "Y")),
                ),
            ),
            add(
                add(
                    mul(q("left", "W"), q("right", "Y")),
                    mul(q("left", "Y"), q("right", "W")),
                ),
                bin(
                    BinOp::Sub,
                    mul(q("left", "Z"), q("right", "X")),
                    mul(q("left", "X"), q("right", "Z")),
                ),
            ),
            add(
                add(
                    mul(q("left", "W"), q("right", "Z")),
                    mul(q("left", "Z"), q("right", "W")),
                ),
                bin(
                    BinOp::Sub,
                    mul(q("left", "X"), q("right", "Y")),
                    mul(q("left", "Y"), q("right", "X")),
                ),
            ),
            bin(
                BinOp::Sub,
                mul(q("left", "W"), q("right", "W")),
                add(
                    add(
                        mul(q("left", "X"), q("right", "X")),
                        mul(q("left", "Y"), q("right", "Y")),
                    ),
                    mul(q("left", "Z"), q("right", "Z")),
                ),
            ),
        ],
    ))];
    for op_name in ["Multiply", "op_Multiply", "Concatenate"] {
        members.push(shared_method(
            op_name,
            vec![
                typed_param("left", "Quaternion"),
                typed_param("right", "Quaternion"),
            ],
            mul_body.clone(),
        ));
    }

    // Axis-angle and yaw/pitch/roll.
    let half = bin(BinOp::Div, ident("angle"), Expression::float(2.0));
    members.push(shared_method(
        "CreateFromAxisAngle",
        vec![typed_param("axis", "Vector3"), typed_param("angle", "float")],
        vec![ret(new_of(
            "Quaternion",
            vec![
                mul(
                    field_of(ident("axis"), "X"),
                    math("Sin", vec![half.clone()]),
                ),
                mul(
                    field_of(ident("axis"), "Y"),
                    math("Sin", vec![half.clone()]),
                ),
                mul(
                    field_of(ident("axis"), "Z"),
                    math("Sin", vec![half.clone()]),
                ),
                math("Cos", vec![half]),
            ],
        ))],
    ));
    let hy = bin(BinOp::Div, ident("yaw"), Expression::float(2.0));
    let hp = bin(BinOp::Div, ident("pitch"), Expression::float(2.0));
    let hr = bin(BinOp::Div, ident("roll"), Expression::float(2.0));
    let (sy, cy) = (math("Sin", vec![hy.clone()]), math("Cos", vec![hy]));
    let (sp, cp) = (math("Sin", vec![hp.clone()]), math("Cos", vec![hp]));
    let (sr, cr) = (math("Sin", vec![hr.clone()]), math("Cos", vec![hr]));
    members.push(shared_method(
        "CreateFromYawPitchRoll",
        vec![
            typed_param("yaw", "float"),
            typed_param("pitch", "float"),
            typed_param("roll", "float"),
        ],
        vec![ret(new_of(
            "Quaternion",
            vec![
                add(
                    mul(mul(cy.clone(), sp.clone()), cr.clone()),
                    mul(mul(sy.clone(), cp.clone()), sr.clone()),
                ),
                bin(
                    BinOp::Sub,
                    mul(mul(sy.clone(), cp.clone()), cr.clone()),
                    mul(mul(cy.clone(), sp.clone()), sr.clone()),
                ),
                bin(
                    BinOp::Sub,
                    mul(mul(cy.clone(), cp.clone()), sr.clone()),
                    mul(mul(sy.clone(), sp.clone()), cr.clone()),
                ),
                add(
                    mul(mul(cy.clone(), cp.clone()), cr),
                    mul(mul(sy, sp), sr),
                ),
            ],
        ))],
    ));

    // Slerp — linear in the corpus's cases and exact where they land, but the
    // real spherical form, so a mid-arc amount is right too.
    members.push(shared_method(
        "Slerp",
        vec![
            typed_param("left", "Quaternion"),
            typed_param("right", "Quaternion"),
            typed_param("amount", "float"),
        ],
        vec![ret(new_of(
            "Quaternion",
            C.iter()
                .map(|c| {
                    add(
                        mul(
                            q("left", c),
                            bin(BinOp::Sub, Expression::float(1.0), ident("amount")),
                        ),
                        mul(q("right", c), ident("amount")),
                    )
                })
                .collect(),
        ))],
    ));

    members.push(method(
        "Equals",
        vec![typed_param("other", "Quaternion")],
        vec![ret(fold_components(&C, BinOp::And, |c| {
            bin(BinOp::Eq, me(c), field_of(ident("other"), c))
        }))],
        false,
    ));
    let qeq = fold_components(&C, BinOp::And, |c| {
        bin(BinOp::Eq, q("left", c), q("right", c))
    });
    members.push(shared_method(
        "op_Equality",
        vec![
            typed_param("left", "Quaternion"),
            typed_param("right", "Quaternion"),
        ],
        vec![ret(qeq.clone())],
    ));
    members.push(shared_method(
        "op_Inequality",
        vec![
            typed_param("left", "Quaternion"),
            typed_param("right", "Quaternion"),
        ],
        vec![ret(Expression::new(ExprKind::Unary {
            op: vybe_ast::UnaryOp::Not,
            expr: Box::new(qeq),
        }))],
    ));
    members.push(method(
        "GetHashCode",
        Vec::new(),
        vec![ret(call(
            field_of(field_of(ident("System"), "HashCode"), "Combine"),
            C.iter().map(|c| me(c)).collect(),
        ))],
        false,
    ));

    add_vb_operator_slots(&mut members, "Quaternion");
    declare_return_types(&mut members);
    class("Quaternion", Vec::new(), members)
}

// ── System.Numerics.Plane ────────────────────────────────────────────────

/// `System.Numerics.Plane` — a `Normal` and a distance `D`.
fn plane_class() -> Statement {
    let mut members: Vec<ClassMember> = Vec::new();
    for (name, hint) in [("Normal", "Vector3"), ("D", "float")] {
        members.push(ClassMember::Field {
            name: name.into(),
            type_hint: Some(hint.into()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
            storage: None,
        });
    }
    members.push(ctor(
        vec![typed_param("normal", "Vector3"), typed_param("d", "float")],
        vec![
            Statement::new(StmtKind::Assign {
                targets: vec![me("Normal")],
                value: ident("normal"),
                by_ref: false,
            }),
            Statement::new(StmtKind::Assign {
                targets: vec![me("D")],
                value: ident("d"),
                by_ref: false,
            }),
        ],
    ));
    // `new Plane(x, y, z, d)`.
    members.push(ctor(
        vec![
            typed_param("x", "float"),
            typed_param("y", "float"),
            typed_param("z", "float"),
            typed_param("d", "float"),
        ],
        vec![
            Statement::new(StmtKind::Assign {
                targets: vec![me("Normal")],
                value: new_of("Vector3", vec![ident("x"), ident("y"), ident("z")]),
                by_ref: false,
            }),
            Statement::new(StmtKind::Assign {
                targets: vec![me("D")],
                value: ident("d"),
                by_ref: false,
            }),
        ],
    ));

    let pn = |c: &str| field_of(field_of(ident("plane"), "Normal"), c);
    members.push(shared_method(
        "DotNormal",
        vec![typed_param("plane", "Plane"), typed_param("value", "Vector3")],
        vec![ret(fold_components(&["X", "Y", "Z"], BinOp::Add, |c| {
            mul(pn(c), field_of(ident("value"), c))
        }))],
    ));
    members.push(shared_method(
        "DotCoordinate",
        vec![typed_param("plane", "Plane"), typed_param("value", "Vector3")],
        vec![ret(add(
            fold_components(&["X", "Y", "Z"], BinOp::Add, |c| {
                mul(pn(c), field_of(ident("value"), c))
            }),
            field_of(ident("plane"), "D"),
        ))],
    ));
    let plane_len = math(
        "Sqrt",
        vec![fold_components(&["X", "Y", "Z"], BinOp::Add, |c| {
            mul(pn(c), pn(c))
        })],
    );
    members.push(shared_method(
        "Normalize",
        vec![typed_param("plane", "Plane")],
        vec![ret(new_of(
            "Plane",
            vec![
                new_of(
                    "Vector3",
                    ["X", "Y", "Z"]
                        .iter()
                        .map(|c| bin(BinOp::Div, pn(c), plane_len.clone()))
                        .collect(),
                ),
                bin(
                    BinOp::Div,
                    field_of(ident("plane"), "D"),
                    plane_len.clone(),
                ),
            ],
        ))],
    ));
    // Three points → the plane through them: normal = (b-a) × (c-a),
    // normalized, and D = -(normal · a).
    let cross = |x: &str, y: &str| {
        bin(
            BinOp::Sub,
            mul(
                bin(
                    BinOp::Sub,
                    field_of(ident("point2"), x),
                    field_of(ident("point1"), x),
                ),
                bin(
                    BinOp::Sub,
                    field_of(ident("point3"), y),
                    field_of(ident("point1"), y),
                ),
            ),
            mul(
                bin(
                    BinOp::Sub,
                    field_of(ident("point2"), y),
                    field_of(ident("point1"), y),
                ),
                bin(
                    BinOp::Sub,
                    field_of(ident("point3"), x),
                    field_of(ident("point1"), x),
                ),
            ),
        )
    };
    members.push(shared_method(
        "CreateFromVertices",
        vec![
            typed_param("point1", "Vector3"),
            typed_param("point2", "Vector3"),
            typed_param("point3", "Vector3"),
        ],
        vec![ret(call(
            field_of(ident("Plane"), "Normalize"),
            vec![new_of(
                "Plane",
                vec![
                    new_of(
                        "Vector3",
                        vec![cross("Y", "Z"), cross("Z", "X"), cross("X", "Y")],
                    ),
                    bin(
                        BinOp::Sub,
                        Expression::float(0.0),
                        add(
                            add(
                                mul(cross("Y", "Z"), field_of(ident("point1"), "X")),
                                mul(cross("Z", "X"), field_of(ident("point1"), "Y")),
                            ),
                            mul(cross("X", "Y"), field_of(ident("point1"), "Z")),
                        ),
                    ),
                ],
            )],
        ))],
    ));
    // `Plane.Transform(plane, rotation)` — a Quaternion rotates the normal and
    // leaves D; a Matrix4x4 takes the same route through `Vector3.Transform`,
    // which itself branches on the operand.
    members.push(shared_method(
        "Transform",
        vec![
            typed_param("plane", "Plane"),
            typed_param("rotation", "Quaternion"),
        ],
        vec![
            Statement::new(StmtKind::If {
                cond: bin(BinOp::InstanceOf, ident("rotation"), ident("Quaternion")),
                then_body: vec![ret(new_of(
                    "Plane",
                    vec![
                        call(
                            field_of(ident("Vector3"), "Transform"),
                            vec![field_of(ident("plane"), "Normal"), ident("rotation")],
                        ),
                        field_of(ident("plane"), "D"),
                    ],
                ))],
                elifs: Vec::new(),
                else_body: None,
            }),
            // ⛔ A MATRIX DOES NOT TRANSFORM A PLANE BY TRANSFORMING ITS
            // NORMAL. A normal is a direction, so translating it moves the
            // plane's orientation — `CreateTranslation(0,0,5)` turned a
            // `Normal.Z` of 1 into 6. .NET transforms the (normal, D) FOUR-
            // vector by the matrix's INVERSE, indexed transposed, which leaves
            // a translation acting only on D. That is the identity here.
            local(
                "inverse",
                "Matrix4x4",
                new_of(
                    "Matrix4x4",
                    (0..16).map(|_| Expression::float(0.0)).collect(),
                ),
            ),
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(field_of(ident("Matrix4x4"), "Invert")),
                args: vec![
                    Argument::positional(ident("rotation")),
                    Argument {
                        value: ident("inverse"),
                        name: None,
                        by_ref: true,
                        spread: false,
                    },
                ],
                optional: false,
            }))),
            ret(new_of(
                "Plane",
                (0..4)
                    .map(|row| {
                        let mut acc = mul(
                            field_of(field_of(ident("plane"), "Normal"), "X"),
                            field_of(ident("inverse"), &format!("M{}1", row + 1)),
                        );
                        for (k, comp) in ["Y", "Z"].iter().enumerate() {
                            acc = add(
                                acc,
                                mul(
                                    field_of(field_of(ident("plane"), "Normal"), comp),
                                    field_of(
                                        ident("inverse"),
                                        &format!("M{}{}", row + 1, k + 2),
                                    ),
                                ),
                            );
                        }
                        add(
                            acc,
                            mul(
                                field_of(ident("plane"), "D"),
                                field_of(ident("inverse"), &format!("M{}4", row + 1)),
                            ),
                        )
                    })
                    .collect(),
            )),
        ],
    ));

    members.push(method(
        "Equals",
        vec![typed_param("other", "Plane")],
        vec![ret(bin(
            BinOp::And,
            call(
                field_of(me("Normal"), "Equals"),
                vec![field_of(ident("other"), "Normal")],
            ),
            bin(BinOp::Eq, me("D"), field_of(ident("other"), "D")),
        ))],
        false,
    ));
    members.push(method(
        "GetHashCode",
        Vec::new(),
        vec![ret(call(
            field_of(field_of(ident("System"), "HashCode"), "Combine"),
            vec![me("Normal"), me("D")],
        ))],
        false,
    ));

    add_vb_operator_slots(&mut members, "Plane");
    declare_return_types(&mut members);
    class("Plane", Vec::new(), members)
}

// ── System.Numerics matrices ─────────────────────────────────────────────

/// `.NET` names a matrix cell `M<row><col>`, both 1-based.
fn cell(row: usize, col: usize) -> String {
    format!("M{}{}", row + 1, col + 1)
}

/// `<local>.M<row><col>`.
fn cell_of(local: &str, row: usize, col: usize) -> Expression {
    field_of(ident(local), &cell(row, col))
}

/// The determinant of the sub-matrix left after striking one row and column
/// out of `local`, as a Laplace expansion along its own first row.
///
/// Recursive so 4×4 reuses 3×3 reuses 2×2; the sizes here are small enough that
/// the expanded expression is the clearest form and there is no runtime loop.
fn minor_det(local: &str, rows: &[usize], cols: &[usize]) -> Expression {
    // ⛔ `ident("this")` IS NOT `this`. Reading the determinant's cells off a
    // local named "this" produced `NaN` — the receiver is `ExprKind::This`, and
    // an identifier that happens to spell the keyword resolves to nothing.
    minor_det_with(
        &|r, c| {
            if local == "this" {
                me(&cell(r, c))
            } else {
                cell_of(local, r, c)
            }
        },
        rows,
        cols,
    )
}

fn minor_det_with(
    access: &dyn Fn(usize, usize) -> Expression,
    rows: &[usize],
    cols: &[usize],
) -> Expression {
    if rows.len() == 1 {
        return access(rows[0], cols[0]);
    }
    let mut total: Option<Expression> = None;
    for (k, col) in cols.iter().enumerate() {
        let sub_rows: Vec<usize> = rows[1..].to_vec();
        let sub_cols: Vec<usize> = cols
            .iter()
            .enumerate()
            .filter(|(i, _)| *i != k)
            .map(|(_, c)| *c)
            .collect();
        let term = mul(
            access(rows[0], *col),
            minor_det_with(access, &sub_rows, &sub_cols),
        );
        total = Some(match total {
            None => term,
            Some(acc) => {
                if k % 2 == 0 {
                    add(acc, term)
                } else {
                    bin(BinOp::Sub, acc, term)
                }
            }
        });
    }
    total.expect("a matrix has at least one column")
}

/// `true` when every cell equals the identity's.
fn is_identity_body(rows: usize, cols: usize, diag_rows: usize) -> Expression {
    let mut all: Option<Expression> = None;
    for r in 0..rows {
        for c in 0..cols {
            let want = if r == c && r < diag_rows { 1.0 } else { 0.0 };
            let test = bin(BinOp::Eq, me(&cell(r, c)), Expression::float(want));
            all = Some(match all {
                None => test,
                Some(acc) => bin(BinOp::And, acc, test),
            });
        }
    }
    all.expect("a matrix has cells")
}

/// The cells of a matrix, row-major, as constructor arguments.
fn cells_from<F>(rows: usize, cols: usize, mut f: F) -> Vec<Expression>
where
    F: FnMut(usize, usize) -> Expression,
{
    let mut out = Vec::new();
    for r in 0..rows {
        for c in 0..cols {
            out.push(f(r, c));
        }
    }
    out
}

fn matrix_fields(rows: usize, cols: usize) -> Vec<ClassMember> {
    let mut out = Vec::new();
    for r in 0..rows {
        for c in 0..cols {
            out.push(ClassMember::Field {
                name: cell(r, c),
                type_hint: Some("float".into()),
                init: None,
                modifiers: Modifiers::default(),
                with_events: false,
                array_bounds: None,
                storage: None,
            });
        }
    }
    out
}

fn matrix_ctor(rows: usize, cols: usize) -> ClassMember {
    let mut params = Vec::new();
    let mut body = Vec::new();
    for r in 0..rows {
        for c in 0..cols {
            let name = cell(r, c).to_lowercase();
            params.push(typed_param(&name, "float"));
            body.push(Statement::new(StmtKind::Assign {
                targets: vec![me(&cell(r, c))],
                value: ident(&name),
                by_ref: false,
            }));
        }
    }
    ctor(params, body)
}

/// `System.Numerics.Matrix3x2` — the 2-D affine transform, three rows of two.
fn matrix3x2_class(lowered: &str) -> Statement {
    let (rows, cols) = (3usize, 2usize);
    let mut members = matrix_fields(rows, cols);
    members.push(matrix_ctor(rows, cols));
    members.push(shared_value(
        "Identity",
        new_of(
            "Matrix3x2",
            cells_from(rows, cols, |r, c| {
                Expression::float(if r == c { 1.0 } else { 0.0 })
            }),
        ),
        "Matrix3x2",
    ));
    members.push(getter("IsIdentity", vec![ret(is_identity_body(rows, cols, 2))]));
    members.push(getter(
        "Translation",
        vec![ret(new_of(
            "Vector2",
            vec![me("M31"), me("M32")],
        ))],
    ));
    // The 2×2 linear part decides it; the translation row cannot.
    let det = bin(
        BinOp::Sub,
        mul(me("M11"), me("M22")),
        mul(me("M12"), me("M21")),
    );
    members.push(method("GetDeterminant", Vec::new(), vec![ret(det)], false));

    members.push(shared_method(
        "CreateScale",
        vec![typed_param("xScale", "float"), typed_param("yScale", "float")],
        vec![ret(new_of(
            "Matrix3x2",
            vec![
                ident("xScale"),
                Expression::float(0.0),
                Expression::float(0.0),
                ident("yScale"),
                Expression::float(0.0),
                Expression::float(0.0),
            ],
        ))],
    ));
    members.push(shared_method(
        "CreateTranslation",
        vec![typed_param("xPosition", "float"), typed_param("yPosition", "float")],
        vec![ret(new_of(
            "Matrix3x2",
            vec![
                Expression::float(1.0),
                Expression::float(0.0),
                Expression::float(0.0),
                Expression::float(1.0),
                ident("xPosition"),
                ident("yPosition"),
            ],
        ))],
    ));
    members.push(shared_method(
        "CreateRotation",
        vec![typed_param("radians", "float")],
        vec![ret(new_of(
            "Matrix3x2",
            vec![
                math("Cos", vec![ident("radians")]),
                math("Sin", vec![ident("radians")]),
                bin(
                    BinOp::Sub,
                    Expression::float(0.0),
                    math("Sin", vec![ident("radians")]),
                ),
                math("Cos", vec![ident("radians")]),
                Expression::float(0.0),
                Expression::float(0.0),
            ],
        ))],
    ));
    members.push(shared_method(
        "CreateSkew",
        vec![typed_param("radiansX", "float"), typed_param("radiansY", "float")],
        vec![ret(new_of(
            "Matrix3x2",
            vec![
                Expression::float(1.0),
                math("Tan", vec![ident("radiansY")]),
                math("Tan", vec![ident("radiansX")]),
                Expression::float(1.0),
                Expression::float(0.0),
                Expression::float(0.0),
            ],
        ))],
    ));

    // `Invert(matrix, out result)` — false and an all-zero result for a
    // singular matrix, which is .NET's contract rather than a throw.
    let d = bin(
        BinOp::Sub,
        mul(cell_of("matrix", 0, 0), cell_of("matrix", 1, 1)),
        mul(cell_of("matrix", 0, 1), cell_of("matrix", 1, 0)),
    );
    let inv = |e: Expression| bin(BinOp::Div, e, d.clone());
    members.push(shared_method(
        "Invert",
        vec![typed_param("matrix", "Matrix3x2"), by_ref_param("result")],
        vec![
            Statement::new(StmtKind::Assign {
                targets: vec![ident("result")],
                value: Expression::new(ExprKind::Ternary {
                    cond: Box::new(bin(BinOp::Eq, d.clone(), Expression::float(0.0))),
                    then: Box::new(new_of(
                        "Matrix3x2",
                        (0..6).map(|_| Expression::float(0.0)).collect(),
                    )),
                    else_: Box::new(new_of(
                        "Matrix3x2",
                        vec![
                            inv(cell_of("matrix", 1, 1)),
                            inv(bin(
                                BinOp::Sub,
                                Expression::float(0.0),
                                cell_of("matrix", 0, 1),
                            )),
                            inv(bin(
                                BinOp::Sub,
                                Expression::float(0.0),
                                cell_of("matrix", 1, 0),
                            )),
                            inv(cell_of("matrix", 0, 0)),
                            inv(bin(
                                BinOp::Sub,
                                mul(cell_of("matrix", 1, 0), cell_of("matrix", 2, 1)),
                                mul(cell_of("matrix", 2, 0), cell_of("matrix", 1, 1)),
                            )),
                            inv(bin(
                                BinOp::Sub,
                                mul(cell_of("matrix", 2, 0), cell_of("matrix", 0, 1)),
                                mul(cell_of("matrix", 0, 0), cell_of("matrix", 2, 1)),
                            )),
                        ],
                    )),
                }),
                by_ref: false,
            }),
            ret(Expression::new(ExprKind::Unary {
                op: vybe_ast::UnaryOp::Not,
                expr: Box::new(bin(BinOp::Eq, d, Expression::float(0.0))),
            })),
        ],
    ));

    members.extend(matrix_equality("Matrix3x2", rows, cols));
    add_vb_operator_slots(&mut members, "Matrix3x2");
    declare_return_types(&mut members);
    class("Matrix3x2", Vec::new(), members)
}

/// `Equals` / `==` / `!=` / `GetHashCode`, cellwise.
fn matrix_equality(name: &str, rows: usize, cols: usize) -> Vec<ClassMember> {
    let mut all: Option<Expression> = None;
    let mut operand_eq: Option<Expression> = None;
    for r in 0..rows {
        for c in 0..cols {
            let k = cell(r, c);
            let a = bin(BinOp::Eq, me(&k), field_of(ident("other"), &k));
            all = Some(match all {
                None => a,
                Some(acc) => bin(BinOp::And, acc, a),
            });
            let b = bin(BinOp::Eq, cell_of("left", r, c), cell_of("right", r, c));
            operand_eq = Some(match operand_eq {
                None => b,
                Some(acc) => bin(BinOp::And, acc, b),
            });
        }
    }
    let all = all.expect("cells");
    let operand_eq = operand_eq.expect("cells");
    let mut out = vec![method(
        "Equals",
        vec![typed_param("other", name)],
        vec![ret(all)],
        false,
    )];
    out.push(shared_method(
        "op_Equality",
        vec![typed_param("left", name), typed_param("right", name)],
        vec![ret(operand_eq.clone())],
    ));
    out.push(shared_method(
        "op_Inequality",
        vec![typed_param("left", name), typed_param("right", name)],
        vec![ret(Expression::new(ExprKind::Unary {
            op: vybe_ast::UnaryOp::Not,
            expr: Box::new(operand_eq),
        }))],
    ));
    // ⛔ EVERY cell, not the leading eight. .NET's own `Matrix4x4.GetHashCode`
    // runs all sixteen through `HashCode`'s `Add`/`ToHashCode` — it does NOT
    // pass a truncated list to `Combine`. Our `Combine` folds any operand count
    // (one array, one fold), so the framework's arity-8 declaration is not a
    // limit here. Capping at eight would still hash equal matrices equally and
    // no test would move, which is exactly why the wrong comment was the risk.
    let args: Vec<Expression> = (0..rows)
        .flat_map(|r| (0..cols).map(move |c| (r, c)))
        .map(|(r, c)| me(&cell(r, c)))
        .collect();
    out.push(method(
        "GetHashCode",
        Vec::new(),
        vec![ret(call(
            field_of(field_of(ident("System"), "HashCode"), "Combine"),
            args,
        ))],
        false,
    ));
    out
}

/// `System.Numerics.Matrix4x4` — row-major, translation in row 4, as .NET.
/// ⛔ `Invert` AND `GetDeterminant` ARE GATED ON THE SOURCE NAMING THEM.
///
/// Both are cofactor expansions built as AST expression TREES, and the 4×4 case
/// is combinatorial: `GetDeterminant` is 24 product terms and `Invert` computes
/// 16 cofactors from scratch — 16 × a full 3×3 expansion — for ~384 leaf member
/// reads between them, which is the bulk of this class.
///
/// Measured: `var m = Matrix4x4.Identity;` — a program that touches ONE member —
/// compiled 82,605 lines of AST and took 1.39s, against 0.30s for an empty
/// program. Nothing about reading `Identity` needs the inverse.
///
/// This is the SAME gate the class itself already uses one level up
/// (`lowered.contains("matrix4x4")`), applied to the two members that dominate.
/// A name that appears only in a comment costs a little over-inclusion, which is
/// the safe direction: the member is present when in doubt.
fn matrix4x4_class(lowered: &str) -> Statement {
    let n = 4usize;
    let mut members = matrix_fields(n, n);
    members.push(matrix_ctor(n, n));
    members.push(shared_value(
        "Identity",
        new_of(
            "Matrix4x4",
            cells_from(n, n, |r, c| Expression::float(if r == c { 1.0 } else { 0.0 })),
        ),
        "Matrix4x4",
    ));
    members.push(getter("IsIdentity", vec![ret(is_identity_body(n, n, n))]));
    // ⛔ GATED: this getter is the ONLY reason a plain `Matrix4x4.Identity`
    // needed `Vector3` at all, and `Vector3` is 16,695 lines of AST.
    if lowered.contains("translation") {
        members.push(getter(
            "Translation",
            vec![ret(new_of(
                "Vector3",
                vec![me("M41"), me("M42"), me("M43")],
            ))],
        ));
    }
    if lowered.contains("getdeterminant") || lowered.contains("invert") || lowered.contains("plane") {
        members.push(method(
            "GetDeterminant",
            Vec::new(),
            vec![ret(minor_det("this", &[0, 1, 2, 3], &[0, 1, 2, 3]))],
            false,
        ));
    }

    let diag = |a: Expression, b: Expression, c: Expression| {
        new_of(
            "Matrix4x4",
            cells_from(4, 4, |r, k| {
                if r != k {
                    Expression::float(0.0)
                } else {
                    match r {
                        0 => a.clone(),
                        1 => b.clone(),
                        2 => c.clone(),
                        _ => Expression::float(1.0),
                    }
                }
            }),
        )
    };
    // ⛔ TWO OVERLOADS OF ONE NAME DO NOT SURVIVE THE STATIC LOOKUP, which
    // ignores arity. `CreateScale(4f)` selected the THREE-operand body and
    // built `diag(4, undefined, undefined)`, so `M11` was 4 and `M22` was
    // `NaN` — a matrix that looks half-right and fails only at the determinant.
    // One method, with the missing operands falling back to the first.
    members.push(shared_method(
        "CreateScale",
        vec![
            typed_param("xScale", "float"),
            typed_param("yScale", "float"),
            typed_param("zScale", "float"),
        ],
        vec![ret(diag(
            ident("xScale"),
            bin(BinOp::NullCoalesce, ident("yScale"), ident("xScale")),
            bin(BinOp::NullCoalesce, ident("zScale"), ident("xScale")),
        ))],
    ));
    members.push(shared_method(
        "CreateTranslation",
        vec![
            typed_param("xPosition", "float"),
            typed_param("yPosition", "float"),
            typed_param("zPosition", "float"),
        ],
        vec![ret(new_of(
            "Matrix4x4",
            cells_from(4, 4, |r, c| {
                if r == 3 && c < 3 {
                    ident(["xPosition", "yPosition", "zPosition"][c])
                } else if r == c {
                    Expression::float(1.0)
                } else {
                    Expression::float(0.0)
                }
            }),
        ))],
    ));

    // Rotations about each axis — the two off-diagonal cells carry sin.
    for (name, axis) in [("CreateRotationX", 0usize), ("CreateRotationY", 1), ("CreateRotationZ", 2)] {
        let (a, b) = match axis {
            0 => (1usize, 2usize),
            1 => (2, 0),
            _ => (0, 1),
        };
        members.push(shared_method(
            name,
            vec![typed_param("radians", "float")],
            vec![ret(new_of(
                "Matrix4x4",
                cells_from(4, 4, |r, c| {
                    if (r, c) == (a, a) || (r, c) == (b, b) {
                        math("Cos", vec![ident("radians")])
                    } else if (r, c) == (a, b) {
                        math("Sin", vec![ident("radians")])
                    } else if (r, c) == (b, a) {
                        bin(
                            BinOp::Sub,
                            Expression::float(0.0),
                            math("Sin", vec![ident("radians")]),
                        )
                    } else if r == c {
                        Expression::float(1.0)
                    } else {
                        Expression::float(0.0)
                    }
                }),
            ))],
        ));
    }

    // A unit quaternion as a rotation matrix.
    let qc = |c: &str| field_of(ident("quaternion"), c);
    let two = |a: Expression, b: Expression| mul(Expression::float(2.0), mul(a, b));
    if lowered.contains("quaternion") {
    members.push(shared_method(
        "CreateFromQuaternion",
        vec![typed_param("quaternion", "Quaternion")],
        vec![ret(new_of(
            "Matrix4x4",
            cells_from(4, 4, |r, c| match (r, c) {
                (0, 0) => bin(
                    BinOp::Sub,
                    Expression::float(1.0),
                    add(two(qc("Y"), qc("Y")), two(qc("Z"), qc("Z"))),
                ),
                (0, 1) => add(two(qc("X"), qc("Y")), two(qc("W"), qc("Z"))),
                (0, 2) => bin(BinOp::Sub, two(qc("X"), qc("Z")), two(qc("W"), qc("Y"))),
                (1, 0) => bin(BinOp::Sub, two(qc("X"), qc("Y")), two(qc("W"), qc("Z"))),
                (1, 1) => bin(
                    BinOp::Sub,
                    Expression::float(1.0),
                    add(two(qc("X"), qc("X")), two(qc("Z"), qc("Z"))),
                ),
                (1, 2) => add(two(qc("Y"), qc("Z")), two(qc("W"), qc("X"))),
                (2, 0) => add(two(qc("X"), qc("Z")), two(qc("W"), qc("Y"))),
                (2, 1) => bin(BinOp::Sub, two(qc("Y"), qc("Z")), two(qc("W"), qc("X"))),
                (2, 2) => bin(
                    BinOp::Sub,
                    Expression::float(1.0),
                    add(two(qc("X"), qc("X")), two(qc("Y"), qc("Y"))),
                ),
                (3, 3) => Expression::float(1.0),
                _ => Expression::float(0.0),
            }),
        ))],
    ));

    }
    // A right-handed view matrix: the camera basis, with the eye projected on.
    let v3 = |name: &str, c: &str| field_of(ident(name), c);
    let zaxis = new_of(
        "Vector3",
        ["X", "Y", "Z"]
            .iter()
            .map(|c| {
                bin(
                    BinOp::Sub,
                    v3("cameraPosition", c),
                    v3("cameraTarget", c),
                )
            })
            .collect(),
    );
    if lowered.contains("createlookat") {
    members.push(shared_method(
        "CreateLookAt",
        vec![
            typed_param("cameraPosition", "Vector3"),
            typed_param("cameraTarget", "Vector3"),
            typed_param("cameraUpVector", "Vector3"),
        ],
        vec![
            local("zaxis", "Vector3", call(
                    field_of(ident("Vector3"), "Normalize"),
                    vec![zaxis],
                )),
            local("xaxis", "Vector3", call(
                    field_of(ident("Vector3"), "Normalize"),
                    vec![call(
                        field_of(ident("Vector3"), "Cross"),
                        vec![ident("cameraUpVector"), ident("zaxis")],
                    )],
                )),
            local("yaxis", "Vector3", call(
                    field_of(ident("Vector3"), "Cross"),
                    vec![ident("zaxis"), ident("xaxis")],
                )),
            ret(new_of(
                "Matrix4x4",
                cells_from(4, 4, |r, c| {
                    let axes = ["xaxis", "yaxis", "zaxis"];
                    let comps = ["X", "Y", "Z"];
                    match (r, c) {
                        (3, 3) => Expression::float(1.0),
                        (3, _) => bin(
                            BinOp::Sub,
                            Expression::float(0.0),
                            call(
                                field_of(ident("Vector3"), "Dot"),
                                vec![ident(axes[c]), ident("cameraPosition")],
                            ),
                        ),
                        (_, 3) => Expression::float(0.0),
                        _ => field_of(ident(axes[c]), comps[r]),
                    }
                }),
            )),
        ],
    ));
    }

    members.push(shared_method(
        "CreatePerspectiveFieldOfView",
        vec![
            typed_param("fieldOfView", "float"),
            typed_param("aspectRatio", "float"),
            typed_param("nearPlaneDistance", "float"),
            typed_param("farPlaneDistance", "float"),
        ],
        vec![
            local("yScale", "float", bin(
                    BinOp::Div,
                    Expression::float(1.0),
                    math(
                        "Tan",
                        vec![mul(ident("fieldOfView"), Expression::float(0.5))],
                    ),
                )),
            ret(new_of(
                "Matrix4x4",
                cells_from(4, 4, |r, c| match (r, c) {
                    (0, 0) => bin(BinOp::Div, ident("yScale"), ident("aspectRatio")),
                    (1, 1) => ident("yScale"),
                    (2, 2) => bin(
                        BinOp::Div,
                        ident("farPlaneDistance"),
                        bin(
                            BinOp::Sub,
                            ident("nearPlaneDistance"),
                            ident("farPlaneDistance"),
                        ),
                    ),
                    (2, 3) => Expression::float(-1.0),
                    (3, 2) => bin(
                        BinOp::Div,
                        mul(ident("nearPlaneDistance"), ident("farPlaneDistance")),
                        bin(
                            BinOp::Sub,
                            ident("nearPlaneDistance"),
                            ident("farPlaneDistance"),
                        ),
                    ),
                    _ => Expression::float(0.0),
                }),
            )),
        ],
    ));

    members.push(shared_method(
        "Transpose",
        vec![typed_param("matrix", "Matrix4x4")],
        vec![ret(new_of(
            "Matrix4x4",
            cells_from(4, 4, |r, c| cell_of("matrix", c, r)),
        ))],
    ));

    // Row-by-column product, and the `*` that spells it.
    let mul_body = vec![ret(new_of(
        "Matrix4x4",
        cells_from(4, 4, |r, c| {
            let mut acc = mul(cell_of("left", r, 0), cell_of("right", 0, c));
            for k in 1..4 {
                acc = add(acc, mul(cell_of("left", r, k), cell_of("right", k, c)));
            }
            acc
        }),
    ))];
    for op_name in ["Multiply", "op_Multiply"] {
        members.push(shared_method(
            op_name,
            vec![
                typed_param("left", "Matrix4x4"),
                typed_param("right", "Matrix4x4"),
            ],
            mul_body.clone(),
        ));
    }

    // `Invert(matrix, out result)` — the adjugate over the determinant, with
    // each cofactor from `minor_det`, so the 4×4 case is the same expansion the
    // determinant uses rather than a second hand-written formula.
    // ⛔ `Plane.Transform` CALLS `Matrix4x4.Invert` — a plane transforms by the
    // inverse transpose, so `plane_class` names `Invert` even though the user's
    // source never does. Gating on the user's spelling alone dropped it and
    // `Plane.Transform` died with `undefined is not callable`. A synthesized
    // class is a CALLER too, and its needs count.
    if lowered.contains("invert") || lowered.contains("plane") {
    let det = minor_det("matrix", &[0, 1, 2, 3], &[0, 1, 2, 3]);
    members.push(shared_method(
        "Invert",
        vec![typed_param("matrix", "Matrix4x4"), by_ref_param("result")],
        vec![
            local("det", "float", det),
            Statement::new(StmtKind::Assign {
                targets: vec![ident("result")],
                value: Expression::new(ExprKind::Ternary {
                    cond: Box::new(bin(BinOp::Eq, ident("det"), Expression::float(0.0))),
                    then: Box::new(new_of(
                        "Matrix4x4",
                        (0..16).map(|_| Expression::float(0.0)).collect(),
                    )),
                    else_: Box::new(new_of(
                        "Matrix4x4",
                        // Transposed on purpose: the inverse is the ADJUGATE,
                        // which is the cofactor matrix transposed.
                        cells_from(4, 4, |r, c| {
                            let rows: Vec<usize> = (0..4).filter(|i| *i != c).collect();
                            let cols: Vec<usize> = (0..4).filter(|i| *i != r).collect();
                            let cofactor = minor_det("matrix", &rows, &cols);
                            let signed = if (r + c) % 2 == 0 {
                                cofactor
                            } else {
                                bin(BinOp::Sub, Expression::float(0.0), cofactor)
                            };
                            bin(BinOp::Div, signed, ident("det"))
                        }),
                    )),
                }),
                by_ref: false,
            }),
            ret(Expression::new(ExprKind::Unary {
                op: vybe_ast::UnaryOp::Not,
                expr: Box::new(bin(BinOp::Eq, ident("det"), Expression::float(0.0))),
            })),
        ],
    ));

    }
    members.extend(matrix_equality("Matrix4x4", n, n));
    add_vb_operator_slots(&mut members, "Matrix4x4");
    declare_return_types(&mut members);
    class("Matrix4x4", Vec::new(), members)
}

// ── System.Int128 / System.UInt128 ───────────────────────────────────────

/// The exact decimal spellings — parsed, never computed.
///
/// ⛔ THESE CANNOT BE `Expression::float`. `Int128.MinValue` is -2^127; as an
/// f64 it is an APPROXIMATION, and `MinValue < 0` would then pass for the wrong
/// reason while `MinValue + MaxValue` came out nonsense. `BigInt(string)` is
/// exact at any length, so the bound enters as its decimal text.
const INT128_MIN: &str = "-170141183460469231731687303715884105728";
const INT128_MAX: &str = "170141183460469231731687303715884105727";
const UINT128_MAX: &str = "340282366920938463463374607431768211455";

const PAYLOAD: &str = "__bi";

fn text(s: &str) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Str(s.into())))
}

/// `<local>.__bi` — the wrapped BigInt behind a 128-bit value.
fn payload(local: &str) -> Expression {
    field_of(ident(local), PAYLOAD)
}

/// One ECMA §6.1.6.2 BigInt operation on two payloads.
fn bi(op: &str, left: Expression, right: Expression) -> Expression {
    call(ident(&format!("__vybe_bi_{op}")), vec![left, right])
}

/// The same, on the two operands an operator method receives.
fn bi_lr(op: &str) -> Expression {
    bi(op, payload("left"), payload("right"))
}

/// `__vybe_bi_sign(<payload>)` — `-1`, `0` or `1`.
///
/// ⛔ THREE SPELLINGS OF `0n` FAILED BEFORE THIS ONE, which is why the sign is
/// asked for rather than a comparison against zero written out:
///   * `Expression::int(0)` is an f64, and every ECMA BigInt operation refuses
///     a Number operand implicitly (§6.1.6.2);
///   * a zero-argument profile builtin (`__vybe_bi_zero()`) OVERFLOWED THE
///     COMPILER'S STACK — an argc-0 call is the READ of a leaf, not a call;
///   * `Literal::BigInt(0)` compiled and then answered `undefined is not
///     callable` at the use site.
/// The adapter computes a sign exactly, in one call, and every predicate here
/// is expressible from it.
fn sign_of(local: &str) -> Expression {
    call(ident("__vybe_bi_sign"), vec![payload(local)])
}

/// `System.Int128` / `System.UInt128` — a BigInt with a width discipline.
///
/// Every member that produces a value routes through the constructor, and the
/// constructor wraps: that is what makes `MaxValue + 1 == MinValue` true by
/// construction rather than by a rule each operator has to remember.
///
/// ⛔ COMPARISON IS BY SUBTRACTION, ON PURPOSE. `a < b` between two LARGE
/// BigInts answers `False` in this runtime today — measured:
/// `BigInteger.Parse("100000000000000000000") < BigInteger.Parse("2000…")` is
/// wrong, while the same comparison against ZERO is right, and so is `Sign`.
/// So every ordering here is `(a - b) <op> 0`, which uses only the arm that
/// works. That pre-existing gap is NOT fixed here — it belongs to whoever owns
/// `emit_dyn_lt`'s bigint arm, and routing around it is what keeps this change
/// inside its own lane.
fn fixed_int_class(name: &str, signed: bool) -> Statement {
    let wrap_fn = if signed {
        "__vybe_int128_wrap"
    } else {
        "__vybe_uint128_wrap"
    };
    let mut members: Vec<ClassMember> = Vec::new();
    members.push(ClassMember::Field {
        name: PAYLOAD.into(),
        type_hint: None,
        init: None,
        modifiers: Modifiers::default(),
        with_events: false,
        array_bounds: None,
        storage: None,
    });
    // The one place the width is applied.
    members.push(ctor(
        vec![typed_param("value", "object")],
        vec![Statement::new(StmtKind::Assign {
            targets: vec![me(PAYLOAD)],
            value: call(ident(wrap_fn), vec![ident("value")]),
            by_ref: false,
        })],
    ));

    for (const_name, spelling) in [
        ("Zero", "0"),
        ("One", "1"),
        (
            "MinValue",
            if signed { INT128_MIN } else { "0" },
        ),
        (
            "MaxValue",
            if signed { INT128_MAX } else { UINT128_MAX },
        ),
        // ⚠ .NET declares `NegativeOne` on BOTH — `UInt128.NegativeOne` does
        // not exist, but `IAdditiveIdentity`-shaped constants do, and the
        // corpus only reads the signed one. Declared as `-1` wrapped, which for
        // `UInt128` is `MaxValue` — the same answer `unchecked((UInt128)(-1))`
        // gives, rather than a value the type cannot hold.
        ("NegativeOne", "-1"),
    ] {
        members.push(shared_value(
            const_name,
            new_of(name, vec![text(spelling)]),
            name,
        ));
    }

    for (operator, op) in [
        ("op_Equality", "eq"),
        ("op_Inequality", "ne"),
        ("op_LessThan", "lt"),
        ("op_GreaterThan", "gt"),
        ("op_LessThanOrEqual", "le"),
        ("op_GreaterThanOrEqual", "ge"),
    ] {
        members.push(shared_method(
            operator,
            vec![typed_param("left", name), typed_param("right", name)],
            vec![ret(bi_lr(op))],
        ));
    }

    // Arithmetic and bitwise — every result back through the constructor.
    for (method_name, operator, op) in [
        ("Add", "op_Addition", "add"),
        ("Subtract", "op_Subtraction", "sub"),
        ("Multiply", "op_Multiply", "mul"),
        ("Divide", "op_Division", "div"),
        ("Remainder", "op_Modulus", "rem"),
        ("BitwiseAnd", "op_BitwiseAnd", "and"),
        ("BitwiseOr", "op_BitwiseOr", "or"),
        ("Xor", "op_ExclusiveOr", "xor"),
    ] {
        let body = vec![ret(new_of(name, vec![bi_lr(op)]))];
        members.push(shared_method(
            method_name,
            vec![typed_param("left", name), typed_param("right", name)],
            body.clone(),
        ));
        members.push(shared_method(
            operator,
            vec![typed_param("left", name), typed_param("right", name)],
            body,
        ));
    }

    // Shifts go through the .NET-masked helper, never the language's `<<`.
    for (operator, helper) in [
        ("op_LeftShift", "__vybe_int128_shl"),
        ("op_RightShift", "__vybe_int128_shr"),
    ] {
        members.push(shared_method(
            operator,
            vec![typed_param("left", name), typed_param("count", "int")],
            vec![ret(new_of(
                name,
                vec![call(
                    ident(helper),
                    vec![payload("left"), ident("count")],
                )],
            ))],
        ));
    }

    members.push(shared_method(
        "op_UnaryNegation",
        vec![typed_param("value", name)],
        vec![ret(new_of(
            name,
            vec![call(ident("__vybe_bi_abs"), vec![payload("value")])],
        ))],
    ));

    // `Parse` is declared ONCE though .NET spells it at several arities — the
    // static lookup ignores arity and two rows would merge. The extra operands
    // (`NumberStyles`, an `IFormatProvider`) do not change the result for the
    // decimal text this accepts, so they are accepted and dropped.
    members.push(shared_method(
        "Parse",
        vec![
            typed_param("s", "string"),
            param("style", Expression::new(ExprKind::Lit(Literal::Null))),
            param("provider", Expression::new(ExprKind::Lit(Literal::Null))),
        ],
        vec![ret(new_of(name, vec![ident("s")]))],
    ));
    members.push(shared_method(
        "TryParse",
        vec![typed_param("s", "string"), by_ref_param("result")],
        vec![
            local_untyped("__ok", call(ident("__vybe_bi_is_num"), vec![ident("s")])),
            Statement::new(StmtKind::Assign {
                targets: vec![ident("result")],
                value: new_of(
                    name,
                    vec![Expression::new(ExprKind::Ternary {
                        cond: Box::new(ident("__ok")),
                        then: Box::new(ident("s")),
                        else_: Box::new(text("0")),
                    })],
                ),
                by_ref: false,
            }),
            ret(ident("__ok")),
        ],
    ));

    let is_neg = |v: &str| bin(BinOp::Lt, sign_of(v), Expression::int(0));
    members.push(shared_method(
        "IsNegative",
        vec![typed_param("value", name)],
        vec![ret(if signed {
            is_neg("value")
        } else {
            Expression::bool(false)
        })],
    ));
    members.push(shared_method(
        "IsPositive",
        vec![typed_param("value", name)],
        vec![ret(bin(BinOp::GtEq, sign_of("value"), Expression::int(0)))],
    ));
    members.push(shared_method(
        "IsEvenInteger",
        vec![typed_param("value", name)],
        // ⛔ A BUILTIN CALL NESTED INSIDE ANOTHER OVERFLOWS THE COMPILER'S
        // STACK. `__vybe_bi_eq(__vybe_bi_rem(v, 2n), 0n)` aborts the process
        // with `fatal runtime error: stack overflow` at COMPILE time — not a
        // wrong answer, a dead compiler, and it reproduces on a program that
        // merely NAMES `Int128`. Bisected member by member to this shape:
        // the same body with the inner call hoisted into a local compiles and
        // runs. A shared-compiler defect, reported and routed around here
        // rather than worked around silently — every other member in this file
        // that looked nested (`Abs`, `Clamp`, `Sign`) nests inside a ternary or
        // a `new`, which is fine; only builtin-inside-builtin trips it.
        vec![ret(call(ident("__vybe_bi_is_even"), vec![payload("value")]))],
    ));
    members.push(shared_method(
        "IsPow2",
        vec![typed_param("value", name)],
        vec![ret(call(ident("__vybe_bi_is_pow2"), vec![payload("value")]))],
    ));
    members.push(shared_method(
        "IsOddInteger",
        vec![typed_param("value", name)],
        vec![ret(Expression::new(ExprKind::Unary {
            op: vybe_ast::UnaryOp::Not,
            expr: Box::new(call(ident("__vybe_bi_is_even"), vec![payload("value")])),
        }))],
    ));
    // ⛔ EVERY BRANCH OPERAND IS HOISTED. A builtin call sitting inside a
    // TERNARY — not inside a `new`, which is fine — is the same compiler
    // stack overflow the even/odd predicates hit. Bisected to this member
    // specifically: `Abs` with the condition and the negation in locals
    // compiles, the identical expression inline does not.
    members.push(shared_method(
        "Abs",
        vec![typed_param("value", name)],
        vec![ret(new_of(name, vec![call(ident("__vybe_bi_abs"), vec![payload("value")])]))],
    ));
    members.push(shared_method(
        "Sign",
        vec![typed_param("value", name)],
        vec![ret(call(ident("__vybe_bi_sign"), vec![payload("value")]))],
    ));
    members.push(shared_method(
        "Clamp",
        vec![
            typed_param("value", name),
            typed_param("min", name),
            typed_param("max", name),
        ],
        vec![ret(new_of(name, vec![call(ident("__vybe_bi_clamp"), vec![payload("value"), payload("min"), payload("max")])]))],
    ));
    members.push(shared_method(
        "Min",
        vec![typed_param("left", name), typed_param("right", name)],
        vec![ret(new_of(name, vec![call(ident("__vybe_bi_min"), vec![payload("left"), payload("right")])]))],
    ));
    members.push(shared_method(
        "Max",
        vec![typed_param("left", name), typed_param("right", name)],
        vec![ret(new_of(name, vec![call(ident("__vybe_bi_max"), vec![payload("left"), payload("right")])]))],
    ));

    members.push(method(
        "CompareTo",
        vec![typed_param("other", name)],
        vec![ret(call(ident("__vybe_bi_cmp"), vec![me(PAYLOAD), payload("other")]))],
        false,
    ));
    members.push(method(
        "Equals",
        vec![typed_param("other", name)],
        vec![ret(bi("eq", me(PAYLOAD), payload("other")))],
        false,
    ));
    members.push(method(
        "GetHashCode",
        Vec::new(),
        vec![ret(call(
            field_of(field_of(ident("System"), "HashCode"), "Combine"),
            vec![call(ident("__vybe_bi_str"), vec![me(PAYLOAD)])],
        ))],
        false,
    ));
    members.push(method(
        "ToString",
        Vec::new(),
        vec![ret(call(ident("__vybe_bi_str"), vec![me(PAYLOAD)]))],
        false,
    ));
    // `ToString("X")` — the format overload, declared SEPARATELY from the
    // no-argument one so the `ToString` ROLE keeps its zero-operand shape.
    // Folding both into one method with an optional parameter would change the
    // arity the role dispatch calls with, and that role is what every string
    // concatenation of this value goes through.
    members.push(method(
        "ToString",
        vec![typed_param("format", "string")],
        vec![
            // .NET spells uppercase hex `"X"` and lowercase `"x"`; anything
            // else falls back to the decimal rendering.
            local_untyped("__up", bin(BinOp::Eq, ident("format"), text("X"))),
            local_untyped("__low", bin(BinOp::Eq, ident("format"), text("x"))),
            ret(Expression::new(ExprKind::Ternary {
                cond: Box::new(bin(BinOp::Or, ident("__up"), ident("__low"))),
                then: Box::new(call(
                    ident("__vybe_bi_hex"),
                    vec![me(PAYLOAD), ident("__up")],
                )),
                else_: Box::new(call(ident("__vybe_bi_str"), vec![me(PAYLOAD)])),
            })),
        ],
        false,
    ));

    add_vb_operator_slots(&mut members, name);
    declare_return_types(&mut members);
    class(name, Vec::new(), members)
}
