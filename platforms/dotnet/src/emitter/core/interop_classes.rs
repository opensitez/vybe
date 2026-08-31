//! `System.Runtime.InteropServices` base classes, synthesized as REAL classes.
//!
//! ⛔ A `ClassType` in `class_exports()` CANNOT BE INHERITED FROM. `--dump-classes`
//! on `Class H : Inherits SafeHandle` lists every exception class and no
//! `safehandle`, and the derived constructor reaches `undefined is not
//! callable` — the shape a missing class always takes. The exceptions are the
//! standing proof of the working alternative: `synthesize_exception_classes`
//! injects them as `StmtKind::ClassDecl`, which is why `Inherits Exception`
//! works and `Inherits SafeHandle` did not.
//!
//! So this is the same move for the same reason, and everything a hand-written
//! emitter had to fake comes back for free: `MyBase.New` binds because there is
//! a real constructor, and `Dispose` reaches the DERIVED `ReleaseHandle`
//! because a method call on `Me` is ordinary virtual dispatch rather than an
//! `emit_invoke_method` with a hand-rolled `__js_this` save/restore.
//!
//! ⚠ Injection is GATED on the program naming the type. The exceptions inject
//! unconditionally; two more classes in every program would shift typeidx
//! numbering for every language, and the class model is mid-conversion.

use vybe_ast::{
    ClassMember, ClassModifiers, ConstructorInitializerTarget, ExprKind, Expression, Modifiers,
    Param, PassBy, PropertySetter, Span, Statement, StmtKind, UnaryOp, Visibility,
};

/// The raw handle. Lowercase because a case-insensitive front end folds a
/// derived class's bare `handle` read to it.
const HANDLE: &str = "handle";
/// ⛔ `IsClosed` IS THE FIELD, not a property over a differently-spelled one.
/// A case-insensitive front end folds both spellings to the same name, so an
/// `IsClosed` getter reading an `isclosed` field reads ITSELF — the first cut
/// of this file did exactly that and every `Dispose` test answered
/// `Stack overflow`. .NET exposes it get-only; a plain field is this
/// platform's convention and cannot recur.
const IS_CLOSED: &str = "IsClosed";
const OWNS_HANDLE: &str = "__owns_handle";

/// Whether `name` is a class this module injects — the same question
/// [`super::exceptions::is_synthesized_exception_class`] answers for the
/// exception hierarchy, and it exists for the same reason: a caller that
/// qualifies bare type names with an imported namespace must leave these alone,
/// or `Inherits SafeHandle` becomes
/// `Inherits System.Runtime.InteropServices.SafeHandle` — a name no class in
/// the program has.
pub fn is_synthesized_interop_class(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "safehandle"
            | "safehandlezeroorminusoneisinvalid"
            | "criticalfinalizerobject"
            | "gchandle"
            | "gchandletype"
    )
}

/// The interop base classes, or nothing when `source` never names one.
pub fn synthesize_interop_classes(source: &str) -> Vec<Statement> {
    let lowered = source.to_ascii_lowercase();
    let mut out = Vec::new();
    if lowered.contains("safehandle") {
        out.push(critical_finalizer_object_class());
        out.push(safe_handle_class());
        out.push(zero_or_minus_one_class());
    }
    if lowered.contains("gchandle") {
        out.push(gc_handle_type_class());
        out.push(gc_handle_class());
    }
    out
}

// ── AST helpers ─────────────────────────────────────────────────────────

pub(super) fn this() -> Expression {
    Expression::new(ExprKind::This)
}

pub(super) fn me(field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(this()),
        field: field.into(),
        null_safe: false,
    })
}

pub(super) fn assign(field: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![me(field)],
        value,
        by_ref: false,
    })
}

fn assign_ident(name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Ident(name.into()))],
        value,
        by_ref: false,
    })
}

/// A BARE call, deliberately — not `Me.method()`.
///
/// ⛔ `ExprKind::Member { object: This }` does not reach a bound receiver in a
/// synthesized class: the ambient `__js_this` branch runs before scope
/// resolution. A bare identifier call is what a language's own
/// implicit-self pass binds, which is why these classes are injected BEFORE
/// those passes rather than appended after them.
fn call_me(method: &str) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident(method.into()))),
        args: Vec::new(),
        optional: false,
    })
}

/// A `ByRef` parameter — `DangerousAddRef(ByRef success As Boolean)` reports
/// through its argument, so the write has to reach the caller's variable.
pub(super) fn by_ref_param(name: &str) -> Param {
    Param {
        pass_by: PassBy::Ref,
        ..param(name, Expression::bool(false))
    }
}

pub(super) fn param(name: &str, default: Expression) -> Param {
    Param {
        name: name.into(),
        type_hint: None,
        default: Some(default),
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: true,
        is_nullable: true,
    }
}

fn field(name: &str) -> ClassMember {
    typed_field(name, "Object")
}

/// ⛔ A DECLARED FIELD NEEDS A DECLARED TYPE. With `type_hint: None` the class
/// table registered no fields at all (`--dump-classes` printed
/// `fields: (none)`), so `Me.IsAllocated` inside the class's own method read
/// `undefined` while `h.IsAllocated` from outside read the value a write had
/// put there dynamically. `GCHandle.Free` therefore saw `Not undefined` and
/// threw "handle is not allocated" on a live handle.
pub(super) fn typed_field(name: &str, type_hint: &str) -> ClassMember {
    ClassMember::Field {
        name: name.into(),
        type_hint: Some(type_hint.into()),
        init: None,
        modifiers: Modifiers::default(),
        with_events: false,
        array_bounds: None,
        storage: None,
    }
}

pub(super) fn method(name: &str, params: Vec<Param>, body: Vec<Statement>, is_sub: bool) -> ClassMember {
    plain_or_virtual_method(name, params, body, is_sub, false)
}

/// ⛔ VIRTUAL ONLY WHERE A DERIVED CLASS ACTUALLY OVERRIDES. Marking every
/// synthesized method `Overridable` put its body in an accessor-shaped chunk
/// where `Me` does not reach the bound receiver: `GCHandle.Free` read
/// `Me.IsAllocated` as `undefined`, so `Not undefined` was True and a live
/// handle threw "not allocated". `SafeHandle.Dispose` had the same defect and
/// LOOKED correct — `Not undefined` sent it down the release path, which is
/// where it was going anyway.
fn virtual_method(name: &str, params: Vec<Param>, body: Vec<Statement>, is_sub: bool) -> ClassMember {
    plain_or_virtual_method(name, params, body, is_sub, true)
}

fn plain_or_virtual_method(
    name: &str,
    params: Vec<Param>,
    body: Vec<Statement>,
    is_sub: bool,
    is_virtual: bool,
) -> ClassMember {
    ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name: name.into(),
        params,
        return_type: None,
        body,
        modifiers: Modifiers {
            is_virtual,
            ..Modifiers::default()
        },
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub,
    })))
}

/// A `Shared` method — `GCHandle.Alloc` is called on the CLASS.
pub(super) fn shared_method(name: &str, params: Vec<Param>, body: Vec<Statement>) -> ClassMember {
    ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
        name: name.into(),
        params,
        return_type: None,
        body,
        // ⛔ BOTH FLAGS. VB's own parser sets `is_static` AND `is_shared` for
        // every `Shared` member; setting only `is_shared` produced a class
        // whose `Alloc` was reachable from nowhere — `GCHandle.Alloc` answered
        // `undefined is not callable` while instance methods on the same
        // synthesized class worked.
        modifiers: Modifiers {
            is_shared: true,
            is_static: true,
            ..Modifiers::default()
        },
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    })))
}

/// A `Shared` field with a constant initialiser — an enum member, in the only
/// shape a synthesized class needs: `GCHandleType.Pinned` is a static read.
fn shared_const(name: &str, value: i64) -> ClassMember {
    ClassMember::Field {
        name: name.into(),
        type_hint: None,
        init: Some(Expression::int(value)),
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

pub(super) fn ident(name: &str) -> Expression {
    Expression::new(ExprKind::Ident(name.into()))
}

/// `Throw New <exception>(<message>)`.
fn throw_new(exception: &str, message: &str) -> Statement {
    Statement::new(StmtKind::Throw {
        cause: None,
        expr: Some(Expression::new(ExprKind::New {
            class: Box::new(ident(exception)),
            args: vec![vybe_ast::Argument::positional(Expression::string(message))],
        })),
    })
}

pub(super) fn getter(name: &str, body: Vec<Statement>) -> ClassMember {
    ClassMember::Property {
        name: name.into(),
        type_hint: None,
        getter: Some(body),
        setter: None::<PropertySetter>,
        is_auto: false,
        modifiers: Modifiers {
            is_virtual: true,
            ..Modifiers::default()
        },
    }
}

pub(super) fn class(name: &str, parents: Vec<String>, members: Vec<ClassMember>) -> Statement {
    Statement::with_span(
        StmtKind::ClassDecl {
            name: name.into(),
            parents,
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
            decorators: Vec::new(),
        },
        Span::default(),
    )
}

// ── The classes ─────────────────────────────────────────────────────────

/// `System.Runtime.ConstrainedExecution.CriticalFinalizerObject` — the real
/// root of the hierarchy, and empty in .NET too: it exists so that
/// `TypeOf h Is CriticalFinalizerObject` answers True, which is precisely what
/// the corpus asks of it.
fn critical_finalizer_object_class() -> Statement {
    class("CriticalFinalizerObject", Vec::new(), Vec::new())
}

fn safe_handle_class() -> Statement {
    // `New(existingHandle, ownsHandle)`. Both optional, because the corpus
    // spells it `MyBase.New(IntPtr.Zero, ownsHandle:=True)` and, in one place,
    // `MyBase.New(ownsHandle)` — one constructor answers both.
    let ctor = ClassMember::Constructor {
        name: None,
        params: vec![
            param("existingHandle", Expression::int(0)),
            param("ownsHandle", Expression::bool(true)),
        ],
        body: vec![
            assign(HANDLE, Expression::new(ExprKind::Ident("existingHandle".into()))),
            assign(
                OWNS_HANDLE,
                Expression::new(ExprKind::Ident("ownsHandle".into())),
            ),
            assign(IS_CLOSED, Expression::bool(false)),
        ],
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    };

    // `Dispose` releases at most once, and only when it owns the handle.
    //
    // ⛔ `Me.ReleaseHandle()` is the whole reason this is a class: the base
    // must reach the DERIVED override. As a synthesized method call it is
    // ordinary virtual dispatch.
    let dispose_body = vec![Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(me(IS_CLOSED)),
        }),
        then_body: vec![
            Statement::new(StmtKind::If {
                cond: me(OWNS_HANDLE),
                then_body: vec![Statement::new(StmtKind::Expr(call_me(
                    "ReleaseHandle",
                )))],
                elifs: Vec::new(),
                else_body: None,
            }),
            assign(IS_CLOSED, Expression::bool(true)),
        ],
        elifs: Vec::new(),
        else_body: None,
    })];

    class(
        "SafeHandle",
        vec!["CriticalFinalizerObject".into()],
        vec![
            field(HANDLE),
            typed_field(IS_CLOSED, "Boolean"),
            typed_field(OWNS_HANDLE, "Boolean"),
            ctor,
            getter("IsInvalid", vec![ret(Expression::bool(false))]),
            method(
                "SetHandle",
                vec![param("h", Expression::int(0))],
                vec![assign(HANDLE, Expression::new(ExprKind::Ident("h".into())))],
                true,
            ),
            method(
                "DangerousGetHandle",
                Vec::new(),
                vec![ret(me(HANDLE))],
                false,
            ),
            method(
                "SetHandleAsInvalid",
                Vec::new(),
                vec![assign(IS_CLOSED, Expression::bool(true))],
                true,
            ),
            virtual_method("ReleaseHandle", Vec::new(), vec![ret(Expression::bool(true))], false),
            // Reference counting keeps a handle alive across a P/Invoke. We
            // have no unmanaged lifetime to protect, so the contract that
            // matters is the OUT parameter: it reports whether the handle was
            // still usable.
            method(
                "DangerousAddRef",
                vec![by_ref_param("success")],
                // ⛔ A CLOSED handle THROWS rather than reporting False.
                // `.NET` raises `ObjectDisposedException`, and the corpus
                // catches it by that exact type — reporting `success = False`
                // would look like a working call that declined.
                vec![
                    Statement::new(StmtKind::If {
                        cond: me(IS_CLOSED),
                        then_body: vec![Statement::new(StmtKind::Throw {
                            cause: None,
                            expr: Some(Expression::new(ExprKind::New {
                                class: Box::new(Expression::new(ExprKind::Ident(
                                    "ObjectDisposedException".into(),
                                ))),
                                args: vec![vybe_ast::Argument::positional(Expression::string(
                                    "SafeHandle",
                                ))],
                            })),
                        })],
                        elifs: Vec::new(),
                        else_body: None,
                    }),
                    assign_ident("success", Expression::bool(true)),
                ],
                true,
            ),
            method("DangerousRelease", Vec::new(), Vec::new(), true),
            method("Dispose", Vec::new(), dispose_body.clone(), true),
            method("Close", Vec::new(), dispose_body.clone(), true),
            method("Finalize", Vec::new(), dispose_body, true),
            // ⛔ THE FINALISER IS WHY `SafeHandle` EXISTS. A handle nobody
            // disposed must still be released when the collector reaches it —
            // in .NET the finaliser runs `Dispose(false)`, which runs
            // `ReleaseHandle`. Named `Finalize`, so VB's
            // `SpecialMethodKind::Destructor` row binds it to the Destructor
            // slot and `GC.WaitForPendingFinalizers` finds it; C#'s `~`
            // spelling reaches the same slot from the other side.
            //
            // Same body as `Dispose`: the release-once protocol is the same
            // one, and the `isclosed` latch already does what
            // `SuppressFinalize` would.
        ],
    )
}

fn zero_or_minus_one_class() -> Statement {
    // Differs from its parent by supplying the `IsInvalid` `SafeHandle` leaves
    // abstract, so it is a parent link and one property.
    class(
        "SafeHandleZeroOrMinusOneIsInvalid",
        vec!["SafeHandle".into()],
        vec![getter(
            "IsInvalid",
            vec![ret(Expression::new(ExprKind::Binary {
                op: vybe_ast::BinOp::Or,
                left: Box::new(eq(me(HANDLE), Expression::int(0))),
                right: Box::new(eq(me(HANDLE), Expression::int(-1))),
            }))],
        )],
    )
}

fn eq(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: vybe_ast::BinOp::Eq,
        left: Box::new(left),
        right: Box::new(right),
    })
}

pub(super) fn ret(value: Expression) -> Statement {
    Statement::new(StmtKind::Return(Some(value)))
}

// ── System.Runtime.InteropServices.GCHandle ─────────────────────────────

/// `GCHandleType` — a static read (`GCHandleType.Pinned`), so a class of
/// `Shared` constants is the whole requirement. Values are .NET's.
fn gc_handle_type_class() -> Statement {
    class(
        "GCHandleType",
        Vec::new(),
        vec![
            shared_const("Weak", 0),
            shared_const("WeakTrackResurrection", 1),
            shared_const("Normal", 2),
            shared_const("Pinned", 3),
        ],
    )
}

const PINNED_TYPE: i64 = 3;

/// `GCHandle`.
///
/// ⛔ THIS REPLACES A WALKER REWRITE, and that is the point. `GCHandle` used to
/// be an anonymous object literal (`dotnet::lowering::gchandle_expr`) built by
/// statement-shape pattern matching in the VB walker: `handle.Free()` was
/// recognised only as a bare statement, `AddrOfPinnedObject` only as a
/// `VarDecl` initialiser. Everything the patterns did not spell — `=`, `<>`,
/// `GetHashCode`, `CType` to `IntPtr`, `Target` on a weak handle — reached
/// `undefined is not callable`. Nine of twenty tests.
///
/// As a real class it needs none of that machinery: reference equality gives
/// `=` / `<>` / `GetHashCode`, and a `CType` through `IntPtr` already
/// round-trips a reference, so the identity of a handle IS the object.
fn gc_handle_class() -> Statement {
    // `Alloc(value)` / `Alloc(value, type)` — one constructor-shaped Shared
    // method answers both, because the second parameter is optional.
    let alloc = shared_method(
        "Alloc",
        vec![
            param("value", Expression::null()),
            param("handleType", Expression::int(2)),
        ],
        vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![vybe_ast::VarDeclarator {
                    pattern: vybe_ast::BindingPattern::Ident("h".into()),
                    type_hint: None,
                    init: Some(Expression::new(ExprKind::New {
                        class: Box::new(ident("GCHandle")),
                        args: Vec::new(),
                    })),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: vybe_ast::VarDeclKind::Dim,
            }),
            assign_member("h", TARGET, ident("value")),
            assign_member("h", IS_ALLOCATED, Expression::bool(true)),
            assign_member(
                "h",
                PINNED,
                eq(ident("handleType"), Expression::int(PINNED_TYPE)),
            ),
            ret(ident("h")),
        ],
    );

    class(
        "GCHandle",
        Vec::new(),
        vec![
            field(TARGET),
            typed_field(IS_ALLOCATED, "Boolean"),
            typed_field(PINNED, "Boolean"),
            alloc,
            // `Free` on an already-freed handle is an error in .NET, not a
            // no-op — the corpus catches `InvalidOperationException` by type.
            method(
                "Free",
                Vec::new(),
                vec![
                    Statement::new(StmtKind::If {
                        cond: Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(me(IS_ALLOCATED)),
                        }),
                        then_body: vec![throw_new(
                            "InvalidOperationException",
                            "Handle is not allocated",
                        )],
                        elifs: Vec::new(),
                        else_body: None,
                    }),
                    assign(IS_ALLOCATED, Expression::bool(false)),
                ],
                true,
            ),
            // Only a PINNED handle has an address; asking an unpinned one is
            // the documented `InvalidOperationException`.
            method(
                "AddrOfPinnedObject",
                Vec::new(),
                vec![
                    Statement::new(StmtKind::If {
                        cond: Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(me(PINNED)),
                        }),
                        then_body: vec![throw_new(
                            "InvalidOperationException",
                            "Handle is not pinned",
                        )],
                        elifs: Vec::new(),
                        else_body: None,
                    }),
                    ret(vybe_compiler::primitives::pointers::make_carray_ptr(
                        me(TARGET),
                        Expression::int(0),
                    )),
                ],
                false,
            ),
            // The handle IS the reference, so both directions are identity.
            shared_method("ToIntPtr", vec![param("h", Expression::null())], vec![ret(ident("h"))]),
            shared_method("FromIntPtr", vec![param("p", Expression::null())], vec![ret(ident("p"))]),
        ],
    )
}

const TARGET: &str = "Target";
const IS_ALLOCATED: &str = "IsAllocated";
const PINNED: &str = "Pinned";

/// `<local>.<field> = <value>` — the one shape `Alloc` needs that `assign`
/// cannot express, since `assign` always writes through `Me`.
fn assign_member(local: &str, field_name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Member {
            object: Box::new(ident(local)),
            field: field_name.into(),
            null_safe: false,
        })],
        value,
        by_ref: false,
    })
}
