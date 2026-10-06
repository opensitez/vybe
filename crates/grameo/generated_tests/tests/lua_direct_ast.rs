#[path = "support/ast_compare.rs"]
mod compare;
#[path = "support/lua_ast.rs"]
mod direct;
#[path = "support/lua_module_samples.rs"]
mod module_samples;
#[path = "support/lua_expression_samples.rs"]
mod samples;
use grameo::{MatchOptions, builder::Builder, engine::build_program};
use vybe_parser_generated_tests::lua::{Parser, Rule};

#[test]
fn native_pratt_direct_ast_preserves_complete_expression_fields_and_spans() {
    use vybe_parser_generated_tests::lua_pratt::Parser as PrattParser;
    for source in samples::cases() {
        let mut builder = direct::Ast::with_pratt(&source);
        let matched = build_program(
            &PrattParser,
            "expr",
            &source,
            MatchOptions::default(),
            &mut builder,
        );
        let reference = walker::expression_for_conformance(&source);
        match (matched, reference) {
            (Ok(matched), Ok((reference, consumed))) => {
                assert_eq!(matched.consumed, consumed, "{source}");
                let actual = builder.into_expression();
                assert!(
                    std::panic::catch_unwind(|| compare::expression(&actual, &reference)).is_ok(),
                    "{source}: actual={actual:?}; reference={reference:?}"
                );
            }
            (Err(_), Err(_)) => assert!(builder.is_empty(), "{source}"),
            (actual, reference) => panic!("{source}: {actual:?}; {reference:?}"),
        }
    }
    for source in ["(a)+(b)+(c)", "a..b..c", "a~-b ~=", "not -a^b"] {
        let mut baseline = direct::Ast::with_pratt(source);
        let steps = build_program(
            &PrattParser,
            "expr",
            source,
            MatchOptions::default(),
            &mut baseline,
        )
        .unwrap()
        .steps;
        for max_steps in 0..steps {
            let mut builder = direct::Ast::with_pratt(source);
            assert!(
                build_program(
                    &PrattParser,
                    "expr",
                    source,
                    MatchOptions {
                        max_steps,
                        ..MatchOptions::default()
                    },
                    &mut builder
                )
                .is_err()
            );
            assert!(builder.is_empty(), "{source}: budget {max_steps}");
        }
    }
    for source in module_samples::cases() {
        let mut builder = direct::Ast::with_pratt(&source);
        build_program(
            &PrattParser,
            "chunk",
            &source,
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap();
        compare::module(&builder.into_module(), &walker::parse(&source).unwrap());
    }
}

#[test]
fn direct_lua_modules_preserve_policies_body_order_and_rollback() {
    for source in module_samples::cases() {
        let mut builder = direct::Ast::new(&source);
        let matched = build_program(
            &Parser,
            "chunk",
            &source,
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap();
        assert_eq!(matched.consumed, source.len());
        let actual = builder.into_module();
        let reference = walker::parse(&source).unwrap();
        compare::module(&actual, &reference);
        assert_eq!(actual.name, "main");
        assert_eq!(actual.language, vybe_ast::Lang::Lua);
        assert_eq!(
            actual.directives.method_receiver,
            Some(vybe_ast::MethodReceiver::CallSite)
        );
        assert_eq!(
            actual.directives.receiver_binding,
            Some(vybe_ast::ReceiverBinding::UniversalParameter)
        );
        assert_eq!(
            actual.directives.type_resolution,
            Some(vybe_ast::TypeResolution::Hybrid)
        );
        assert_eq!(
            actual.directives.operator_dispatch,
            Some(vybe_ast::OperatorDispatch::RuntimeProtocol)
        );
    }
    let source = "function a.b(x,...) for k,v in pairs(x) do if v then return k,... end end end";
    let mut baseline = direct::Ast::new(source);
    let steps = build_program(
        &Parser,
        "chunk",
        source,
        MatchOptions::default(),
        &mut baseline,
    )
    .unwrap()
    .steps;
    for max_steps in (0..steps).step_by(37) {
        let mut builder = direct::Ast::new(source);
        assert!(
            build_program(
                &Parser,
                "chunk",
                source,
                MatchOptions {
                    max_steps,
                    ..MatchOptions::default()
                },
                &mut builder
            )
            .is_err()
        );
        assert!(builder.is_empty(), "budget {max_steps}");
    }
    for source in [
        "local x=",
        "return 1+",
        "if a then",
        "local x=1 @",
        "return 1 end",
    ] {
        let mut builder = direct::Ast::new(source);
        assert!(
            build_program(
                &Parser,
                "chunk",
                source,
                MatchOptions::default(),
                &mut builder
            )
            .is_err()
        );
        assert!(builder.is_empty(), "{source}");
    }
}
struct LuaParser;

#[test]
fn colon_declarations_insert_self_once_and_preserve_existing_field() {
    use vybe_ast::{ExprKind, LambdaBody, StmtKind};
    for (source, names) in [
        ("function t:m() end", vec!["self"]),
        ("function t:m(self,x) return self,x end", vec!["self", "x"]),
        ("function a.b:m(x,self) end", vec!["self", "x", "self"]),
        ("function t:m(...) end", vec!["self", "_lua_varargs"]),
    ] {
        let mut builder = direct::Ast::new(source);
        build_program(
            &Parser,
            "chunk",
            source,
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap();
        let actual = builder.into_module();
        let StmtKind::Block(statements) = &actual.body[0].kind else {
            panic!("missing method block")
        };
        assert_eq!(statements.len(), 2);
        let StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } = &statements[0].kind
        else {
            panic!("missing data preservation")
        };
        assert_eq!(then_body.len(), 2);
        assert!(elifs.is_empty());
        assert!(else_body.is_none());
        let StmtKind::Assign {
            targets,
            value,
            by_ref,
        } = &statements[1].kind
        else {
            panic!("missing method assignment")
        };
        assert_eq!(targets.len(), 1);
        assert!(!by_ref);
        let ExprKind::Lambda {
            params,
            body: LambdaBody::Block(_),
            ..
        } = &value.kind
        else {
            panic!("missing method lambda")
        };
        assert_eq!(
            params.iter().map(|p| p.name.as_str()).collect::<Vec<_>>(),
            names
        );
        compare::module(&actual, &walker::parse(source).unwrap());
    }
}
impl LuaParser {
    fn parse(
        rule: Rule,
        source: &str,
    ) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
        Parser.parse_pairs(rule, source)
    }
}
#[allow(
    clippy::collapsible_if,
    clippy::while_let_on_iterator,
    clippy::clone_on_copy,
    clippy::manual_strip,
    dead_code
)]
mod walker {
    include!(concat!(env!("OUT_DIR"), "/lua_native_walker.rs"));
}
#[allow(dead_code)]
mod normalize {
    include!(concat!(env!("OUT_DIR"), "/lua_numeric_helpers.rs"));
}
fn module(expression: vybe_ast::Expression) -> vybe_ast::Module {
    vybe_ast::Module {
        name: "expression".into(),
        language: vybe_ast::Lang::Lua,
        body: vec![vybe_ast::Statement::new(vybe_ast::StmtKind::Return(Some(
            expression,
        )))],
        imports: Vec::new(),
        directives: Default::default(),
        canon: Default::default(),
    }
}
fn check(source: &str) {
    let mut builder = direct::Ast::new(source);
    let matched = build_program(
        &Parser,
        "expr",
        source,
        MatchOptions::default(),
        &mut builder,
    );
    let reference = walker::expression_for_conformance(source);
    match (matched, reference) {
        (Ok(matched), Ok((reference, consumed))) => {
            assert_eq!(matched.consumed, consumed, "{source}");
            let direct = builder.into_expression();
            let direct = module(direct);
            let reference = module(reference);
            assert!(
                std::panic::catch_unwind(|| compare::module(&direct, &reference)).is_ok(),
                "{source}: direct={direct:?}; reference={reference:?}"
            );
        }
        (Err(_), Err(_)) => {
            assert_eq!(builder.checkpoint(), 0);
            assert!(builder.is_empty());
        }
        (direct, reference) => panic!("{source}: direct={direct:?}; reference={reference:?}"),
    }
}
#[test]
fn unchanged_lua_grammar_constructs_expression_ast_without_captures_or_walker() {
    for source in samples::cases() {
        check(&source);
    }
}
#[test]
fn expression_builder_restores_frames_and_owned_values_on_resource_or_binding_failure() {
    for source in ["f(1, g(2)).x + 3", "{1, 2}", "function() return 1 end"] {
        for max_steps in [0, 1, 20, 100, 300, 10_000] {
            let mut builder = direct::Ast::new(source);
            let result = build_program(
                &Parser,
                "expr",
                source,
                MatchOptions {
                    max_steps,
                    ..MatchOptions::default()
                },
                &mut builder,
            );
            if result.is_err() {
                assert_eq!(builder.checkpoint(), 0, "{source}: {max_steps}");
                assert!(builder.is_empty(), "{source}: {max_steps}");
            }
        }
    }
}

#[test]
fn table_field_alternatives_restore_owned_keys_fields_and_variant_markers() {
    for source in [
        "{a}",
        "{... + 1}",
        "{[a]=f(), name=2, ...t}",
        "{a=",
        "{[a]=function() return 1 end}",
    ] {
        let mut baseline = direct::Ast::new(source);
        let baseline_result = build_program(
            &Parser,
            "expr",
            source,
            MatchOptions::default(),
            &mut baseline,
        );
        let bound = baseline_result
            .as_ref()
            .map_or(2_000, |result| result.steps + 2);
        if baseline_result.is_err() {
            assert!(baseline.is_empty());
        }
        let mut saw_success = false;
        for max_steps in 0..bound {
            let mut builder = direct::Ast::new(source);
            let result = build_program(
                &Parser,
                "expr",
                source,
                MatchOptions {
                    max_steps,
                    ..MatchOptions::default()
                },
                &mut builder,
            );
            if result.is_err() {
                assert!(builder.is_empty(), "{source}: {max_steps}");
            } else {
                saw_success = true;
            }
        }
        assert_eq!(saw_success, baseline_result.is_ok(), "{source}");
    }
}

#[test]
fn function_bindings_restore_parameters_bodies_and_local_declarations_on_failure() {
    for source in [
        "function(a,...) local x,y = a, f() return x,... end",
        "function() local function f(x) return x end return f(1) end",
        "function() function f(a,...) local x=a return x,... end end",
        "function() function f() end function g(self) return self end end",
        "function() function a.b.c(x,...) return x,... end end",
        "function() function a.b:c(x,...) return self,x,... end end",
        "function() function t:m(self,x) return self,x end end",
        "function() function t:m() end end",
        "function() local a=1, f(), g() return a end",
    ] {
        for max_steps in (0..8_000).step_by(37) {
            let mut builder = direct::Ast::new(source);
            let result = build_program(
                &Parser,
                "expr",
                source,
                MatchOptions {
                    max_steps,
                    ..MatchOptions::default()
                },
                &mut builder,
            );
            if result.is_err() {
                assert!(builder.is_empty(), "{source}: {max_steps}");
            }
        }
    }
}

#[test]
fn assignments_and_scoped_loop_bodies_restore_owned_children_on_failure() {
    for source in [
        "function() a,b={1,2},f() return a,b end",
        "function() while a do local x=1 a=x break end end",
        "function() repeat do a,b=b,a end until f() end",
    ] {
        let mut baseline = direct::Ast::new(source);
        let steps = build_program(
            &Parser,
            "expr",
            source,
            MatchOptions::default(),
            &mut baseline,
        )
        .unwrap()
        .steps;
        let mut succeeded = false;
        for max_steps in (0..steps + 38).step_by(37) {
            let mut builder = direct::Ast::new(source);
            let result = build_program(
                &Parser,
                "expr",
                source,
                MatchOptions {
                    max_steps,
                    ..MatchOptions::default()
                },
                &mut builder,
            );
            if result.is_err() {
                assert!(builder.is_empty(), "{source}: {max_steps}");
            } else {
                succeeded = true;
            }
        }
        assert!(succeeded);
    }
}

#[test]
fn empty_conditional_branches_keep_then_elseif_and_else_roles() {
    use vybe_ast::{ExprKind, LambdaBody, StmtKind};
    for (statement, elif_count, else_len) in [
        ("if false then else return 2 end", 0, Some(1)),
        (
            "if false then elseif true then return 2 else return 3 end",
            1,
            Some(1),
        ),
        ("if false then elseif true then else end", 1, Some(0)),
        ("if false then elseif true then end", 1, None),
    ] {
        let source = format!("function() {statement} end");
        let mut builder = direct::Ast::new(&source);
        let steps = build_program(
            &Parser,
            "expr",
            &source,
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap()
        .steps;
        let ExprKind::Lambda {
            body: LambdaBody::Block(body),
            ..
        } = builder.into_expression().kind
        else {
            panic!("expected lambda")
        };
        let StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } = &body[0].kind
        else {
            panic!("expected conditional")
        };
        assert!(then_body.is_empty(), "{statement}");
        assert_eq!(elifs.len(), elif_count, "{statement}");
        assert_eq!(else_body.as_ref().map(Vec::len), else_len, "{statement}");
        for max_steps in (0..steps).step_by(37) {
            let mut builder = direct::Ast::new(&source);
            let result = build_program(
                &Parser,
                "expr",
                &source,
                MatchOptions {
                    max_steps,
                    ..MatchOptions::default()
                },
                &mut builder,
            );
            assert!(result.is_err());
            assert!(builder.is_empty(), "{statement}: {max_steps}");
        }
    }
}

#[test]
fn generic_for_restores_exact_names_iterators_and_scoped_bodies() {
    for source in [
        "function() for k,v in pairs(t) do f(k,v) end end",
        "function() for a,b__lua_extra__c,d in iter,state,control do local x=d end end",
        "function() for x in (f()) do end end",
    ] {
        let mut baseline = direct::Ast::new(source);
        let steps = build_program(
            &Parser,
            "expr",
            source,
            MatchOptions::default(),
            &mut baseline,
        )
        .unwrap()
        .steps;
        for max_steps in (0..steps).step_by(37) {
            let mut builder = direct::Ast::new(source);
            assert!(
                build_program(
                    &Parser,
                    "expr",
                    source,
                    MatchOptions {
                        max_steps,
                        ..MatchOptions::default()
                    },
                    &mut builder
                )
                .is_err()
            );
            assert!(builder.is_empty(), "{source}: {max_steps}");
        }
    }
}

#[test]
fn numeric_for_preserves_bounds_scope_and_exact_rollback() {
    for source in [
        "for i=1,2 do end",
        "for i=1,2,0 do return i end",
        "for i=9007199254740993,4.5 do t[i]=function() return i,__lua_capture_i end end",
        "for i=1,3 do local f=function() return function(i) return i end,i end end",
        "for i=1,3 do for j=1,2 do t[i]=function() return i,j end end end",
    ] {
        verify(&Parser, source, false);
        verify(
            &vybe_parser_generated_tests::lua_pratt::Parser,
            source,
            true,
        );
    }
    fn verify<P: grameo::program::Program>(program: &P, source: &str, pratt: bool) {
        let mut baseline = if pratt {
            direct::Ast::with_pratt(source)
        } else {
            direct::Ast::new(source)
        };
        let steps = build_program(
            program,
            "chunk",
            source,
            MatchOptions::default(),
            &mut baseline,
        )
        .unwrap()
        .steps;
        let actual = baseline.into_module();
        compare::module(&actual, &walker::parse(source).unwrap());
        let vybe_ast::StmtKind::For {
            init: Some(init),
            body,
            cond: Some(_),
            update: Some(_),
            ..
        } = &actual.body[0].kind
        else {
            panic!("missing numeric for")
        };
        let vybe_ast::StmtKind::VarDecl { declarations, .. } = &init.kind else {
            panic!("missing bound declarations")
        };
        assert_eq!(declarations.len(), 3);
        assert!(matches!(body[0].kind, vybe_ast::StmtKind::Block(_)));
        // Force failure after every operation, including completed lowering
        // followed by chunk EOI/trivia, to exercise composition inverses.
        for max_steps in (0..steps).step_by(17) {
            let mut builder = if pratt {
                direct::Ast::with_pratt(source)
            } else {
                direct::Ast::new(source)
            };
            assert!(
                build_program(
                    program,
                    "chunk",
                    source,
                    MatchOptions {
                        max_steps,
                        ..MatchOptions::default()
                    },
                    &mut builder
                )
                .is_err()
            );
            assert!(
                builder.is_empty(),
                "{source}: pratt={pratt}, budget={max_steps}"
            );
        }
        for invalid in [
            format!("{source} ?"),
            source.trim_end_matches("end").to_owned(),
        ] {
            let mut builder = if pratt {
                direct::Ast::with_pratt(&invalid)
            } else {
                direct::Ast::new(&invalid)
            };
            assert!(
                build_program(
                    program,
                    "chunk",
                    &invalid,
                    MatchOptions::default(),
                    &mut builder
                )
                .is_err()
            );
            assert!(builder.is_empty(), "{invalid}");
        }
    }
}
