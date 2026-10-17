use vybe_ast::*;
use vybe_ast::class_normalize::{base_call_for, base_call_for_owned,
    from_constructor_member, from_constructor_member_owned};

#[test]
fn owned_constructor_matches_borrowed_policy_and_transfers_storage() {
    for target in [ConstructorInitializerTarget::Base, ConstructorInitializerTarget::This] {
        for has_parent in [false, true] {
            for args in [None, Some(vec![]), Some(vec![Expression::ident("base_argument")])] {
                let params = vec![Param {
                    name: "value".into(), type_hint: None, default: Some(Expression::int(7)),
                    pass_by: PassBy::Alias, is_rest: false, is_kwargs: false,
                    is_optional: true, is_nullable: true,
                }];
                let body = vec![Statement::new(StmtKind::Return(Some(Expression::ident("value"))))];
                let params_ptr = params.as_ptr();
                let body_ptr = body.as_ptr();
                let argument_ptr = args.as_ref().and_then(|args| args.first()).map(|expr| {
                    let ExprKind::Ident(name) = &expr.kind else { unreachable!() };
                    name.as_ptr()
                });
                let member = ClassMember::Constructor { name: Some("named".into()), params, body,
                    base_args: args, initializer_target: target, visibility: Visibility::Private };
                let span = Span { start_line: 3, start_col: 4, end_line: 8, end_col: 2 };
                let expected = from_constructor_member(span, &member, has_parent).unwrap();
                let actual = from_constructor_member_owned(span, member, has_parent).unwrap();
                assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                assert_eq!(actual.params.as_ptr(), params_ptr);
                assert_eq!(actual.body.as_ptr(), body_ptr);
                if let Some(pointer) = argument_ptr {
                    let base_args = match &actual.base_call {
                        class_normalize::BaseCall::Explicit(args) | class_normalize::BaseCall::This(args) => args,
                        _ => panic!("initializer arguments lost"),
                    };
                    let ExprKind::Ident(name) = &base_args[0].value.kind else { panic!("argument changed") };
                    assert_eq!(name.as_ptr(), pointer);
                }
            }
        }
    }
}

#[test]
fn owned_base_call_preserves_absent_and_explicit_empty_initializers() {
    for target in [ConstructorInitializerTarget::Base, ConstructorInitializerTarget::This] {
        for has_parent in [false, true] {
            for args in [None, Some(vec![]), Some(vec![Expression::int(2)])] {
                let expected = base_call_for(args.as_deref(), target, has_parent);
                let actual = base_call_for_owned(args, target, has_parent);
                assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
            }
        }
    }
}

#[test]
fn unhandled_owned_member_is_returned_intact() {
    let statement = Box::new(Statement::new(StmtKind::Return(Some(Expression::int(1)))));
    let pointer = statement.as_ref() as *const Statement;
    let result = from_constructor_member_owned(Span::default(), ClassMember::Method(statement), true);
    let ClassMember::Method(statement) = result.unwrap_err() else { panic!("member changed") };
    assert_eq!(statement.as_ref() as *const Statement, pointer);
}
