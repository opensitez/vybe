//! Explicit structural comparator for the raw Lua walker's AST vocabulary.
//! Unknown variants fail rather than silently skipping future language output.
use vybe_ast::*;
fn span(a: Span, b: Span) {
    assert_eq!(
        (a.start_line, a.start_col, a.end_line, a.end_col),
        (b.start_line, b.start_col, b.end_line, b.end_col)
    );
}
fn list<T>(a: &[T], b: &[T], compare: fn(&T, &T)) {
    assert_eq!(a.len(), b.len());
    for (a, b) in a.iter().zip(b) {
        compare(a, b);
    }
}
fn option<T>(a: &Option<T>, b: &Option<T>, compare: fn(&T, &T)) {
    match (a, b) {
        (Some(a), Some(b)) => compare(a, b),
        (None, None) => {}
        _ => panic!("optional AST field differs"),
    }
}
// These metadata types contain no AST subtrees; compare their complete Debug
// representations because several scalar policy structs do not implement Eq.
fn metadata<T: std::fmt::Debug>(a: &T, b: &T) {
    assert_eq!(format!("{a:?}"), format!("{b:?}"));
}
pub fn module(a: &Module, b: &Module) {
    let Module {
        name,
        language,
        body,
        imports,
        directives,
        canon,
    } = a;
    assert_eq!(name, &b.name);
    assert_eq!(language, &b.language);
    assert!(
        imports.is_empty() && b.imports.is_empty(),
        "Lua raw walker should emit no imports"
    );
    assert_eq!(directives, &b.directives);
    assert_eq!(canon, &b.canon);
    list(body, &b.body, statement);
}
fn param(a: &Param, b: &Param) {
    let Param {
        name,
        type_hint,
        default,
        pass_by,
        is_rest,
        is_kwargs,
        is_optional,
        is_nullable,
    } = a;
    assert_eq!(name, &b.name);
    assert_eq!(type_hint, &b.type_hint);
    option(default, &b.default, expression);
    assert_eq!(
        (*pass_by, *is_rest, *is_kwargs, *is_optional, *is_nullable),
        (
            b.pass_by,
            b.is_rest,
            b.is_kwargs,
            b.is_optional,
            b.is_nullable
        )
    );
}
fn declaration(a: &VarDeclarator, b: &VarDeclarator) {
    let VarDeclarator {
        pattern,
        type_hint,
        init,
        array_bounds,
        with_events,
    } = a;
    match (pattern, &b.pattern) {
        (BindingPattern::Ident(a), BindingPattern::Ident(b)) => assert_eq!(a, b),
        _ => panic!("unsupported Lua binding pattern"),
    }
    assert_eq!(type_hint, &b.type_hint);
    option(init, &b.init, expression);
    option(array_bounds, &b.array_bounds, |a, b| list(a, b, expression));
    assert_eq!(with_events, &b.with_events);
}
fn statement(a: &Statement, b: &Statement) {
    span(a.span, b.span);
    use StmtKind::*;
    match (&a.kind, &b.kind) {
        (Expr(a), Expr(b)) => expression(a, b),
        (Block(a), Block(b)) => list(a, b, statement),
        (Return(a), Return(b)) => option(a, b, expression),
        (Break(BreakTarget::Implicit), Break(BreakTarget::Implicit)) => {}
        (GoTo(a), GoTo(b)) | (Label(a), Label(b)) => assert_eq!(a, b),
        (
            Assign {
                targets: a,
                value: av,
                by_ref: ar,
            },
            Assign {
                targets: b,
                value: bv,
                by_ref: br,
            },
        ) => {
            list(a, b, expression);
            expression(av, bv);
            assert_eq!(ar, br);
        }
        (
            VarDecl {
                declarations: a,
                kind: ak,
            },
            VarDecl {
                declarations: b,
                kind: bk,
            },
        ) => {
            assert_eq!(ak, bk);
            list(a, b, declaration);
        }
        (
            FunctionDecl {
                name: a,
                params: ap,
                return_type: at,
                body: ab,
                modifiers: am,
                handles: ah,
                is_async: aa,
                is_generator: ag,
                is_sub: asub,
            },
            FunctionDecl {
                name: b,
                params: bp,
                return_type: bt,
                body: bb,
                modifiers: bm,
                handles: bh,
                is_async: ba,
                is_generator: bg,
                is_sub: bsub,
            },
        ) => {
            assert_eq!(a, b);
            assert_eq!(at, bt);
            assert_eq!(ah, bh);
            assert_eq!((aa, ag, asub), (ba, bg, bsub));
            metadata(am, bm);
            list(ap, bp, param);
            list(ab, bb, statement);
        }
        (
            While {
                cond: a,
                body: ab,
                else_body: ae,
            },
            While {
                cond: b,
                body: bb,
                else_body: be,
            },
        ) => {
            expression(a, b);
            list(ab, bb, statement);
            option(ae, be, |a, b| list(a, b, statement));
        }
        (
            DoWhile {
                cond: a,
                body: ab,
                until: au,
            },
            DoWhile {
                cond: b,
                body: bb,
                until: bu,
            },
        ) => {
            expression(a, b);
            list(ab, bb, statement);
            assert_eq!(au, bu);
        }
        (
            For {
                init: a,
                cond: ac,
                update: au,
                body: ab,
            },
            For {
                init: b,
                cond: bc,
                update: bu,
                body: bb,
            },
        ) => {
            option(a, b, |a, b| statement(a, b));
            option(ac, bc, expression);
            option(au, bu, expression);
            list(ab, bb, statement);
        }
        (
            ForIn {
                var: a,
                key: ak,
                iter: ai,
                body: ab,
                of: ao,
                else_body: ae,
                is_async: aa,
            },
            ForIn {
                var: b,
                key: bk,
                iter: bi,
                body: bb,
                of: bo,
                else_body: be,
                is_async: ba,
            },
        ) => {
            assert_eq!((a, ak, ao, aa), (b, bk, bo, ba));
            expression(ai, bi);
            list(ab, bb, statement);
            option(ae, be, |a, b| list(a, b, statement));
        }
        (
            If {
                cond: a,
                then_body: at,
                elifs: ai,
                else_body: ae,
            },
            If {
                cond: b,
                then_body: bt,
                elifs: bi,
                else_body: be,
            },
        ) => {
            expression(a, b);
            list(at, bt, statement);
            list(ai, bi, |a, b| {
                expression(&a.0, &b.0);
                list(&a.1, &b.1, statement);
            });
            option(ae, be, |a, b| list(a, b, statement));
        }
        (a, b) => panic!("unhandled or different Lua statements: {a:?} / {b:?}"),
    }
}
fn literal(a: &Literal, b: &Literal) {
    use Literal::*;
    match (a, b) {
        (Int(a), Int(b)) | (BigInt(a), BigInt(b)) => assert_eq!(a, b),
        (Float(a), Float(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        (Str(a), Str(b)) => assert_eq!(a, b),
        (Bool(a), Bool(b)) => assert_eq!(a, b),
        (Char(a), Char(b)) => assert_eq!(a, b),
        (Bytes(a), Bytes(b)) => assert_eq!(a, b),
        (Null, Null) | (Undefined, Undefined) | (Ellipsis, Ellipsis) => {}
        _ => panic!("literal kinds differ"),
    }
}
fn expression(a: &Expression, b: &Expression) {
    span(a.span, b.span);
    use ExprKind::*;
    match (&a.kind, &b.kind) {
        (Lit(a), Lit(b)) => literal(a, b),
        (Ident(a), Ident(b)) => assert_eq!(a, b),
        (This, This) | (Super, Super) | (GlobalNamespace, GlobalNamespace) => {}
        (
            Binary {
                op: a,
                left: al,
                right: ar,
            },
            Binary {
                op: b,
                left: bl,
                right: br,
            },
        ) => {
            assert_eq!(a, b);
            expression(al, bl);
            expression(ar, br);
        }
        (Unary { op: a, expr: ae }, Unary { op: b, expr: be }) => {
            assert_eq!(a, b);
            expression(ae, be);
        }
        (
            Ternary {
                cond: a,
                then: at,
                else_: ae,
            },
            Ternary {
                cond: b,
                then: bt,
                else_: be,
            },
        ) => {
            expression(a, b);
            expression(at, bt);
            expression(ae, be);
        }
        (
            Assign {
                target: a,
                value: av,
            },
            Assign {
                target: b,
                value: bv,
            },
        ) => {
            expression(a, b);
            expression(av, bv);
        }
        (
            Member {
                object: a,
                field: af,
                null_safe: an,
            },
            Member {
                object: b,
                field: bf,
                null_safe: bn,
            },
        ) => {
            expression(a, b);
            assert_eq!((af, an), (bf, bn));
        }
        (
            Index {
                object: a,
                index: ai,
                null_safe: an,
            },
            Index {
                object: b,
                index: bi,
                null_safe: bn,
            },
        ) => {
            expression(a, b);
            expression(ai, bi);
            assert_eq!(an, bn);
        }
        (Spread(a), Spread(b)) => expression(a, b),
        (Tuple(a), Tuple(b)) | (Sequence(a), Sequence(b)) => list(a, b, expression),
        (Array(a), Array(b)) => list(a, b, |a, b| {
            let ArrayElement {
                key,
                value,
                spread,
                by_ref,
            } = a;
            option(key, &b.key, expression);
            expression(value, &b.value);
            assert_eq!((spread, by_ref), (&b.spread, &b.by_ref));
        }),
        (
            Call {
                callee: a,
                args: aa,
                optional: ao,
            },
            Call {
                callee: b,
                args: ba,
                optional: bo,
            },
        ) => {
            expression(a, b);
            assert_eq!(ao, bo);
            list(aa, ba, |a, b| {
                let Argument {
                    value,
                    name,
                    by_ref,
                    spread,
                } = a;
                expression(value, &b.value);
                assert_eq!((name, by_ref, spread), (&b.name, &b.by_ref, &b.spread));
            });
        }
        (
            Lambda {
                params: a,
                body: ab,
                is_async: aa,
                captures: ac,
            },
            Lambda {
                params: b,
                body: bb,
                is_async: ba,
                captures: bc,
            },
        ) => {
            list(a, b, param);
            assert_eq!((aa, ac), (ba, bc));
            match (ab, bb) {
                (LambdaBody::Expr(a), LambdaBody::Expr(b)) => expression(a, b),
                (LambdaBody::Block(a), LambdaBody::Block(b)) => list(a, b, statement),
                _ => panic!("lambda body kinds differ"),
            }
        }
        (a, b) => panic!("unhandled or different Lua expressions: {a:?} / {b:?}"),
    }
}
