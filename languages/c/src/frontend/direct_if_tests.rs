use super::*;
use grameo::{CompileOptions, CompiledGrammar, MatchOptions, engine};
use std::sync::LazyLock;

static GRAMMAR: LazyLock<CompiledGrammar> = LazyLock::new(||
    grameo::compile_with_options(include_str!("../grammar.pest"), CompileOptions::default()).unwrap());

fn captures<'a>(pairs: grameo::compat::Pairs<'a, Rule>) -> Vec<Pair<'a, Rule>> {
    pairs.flat_map(|pair| std::iter::once(pair.clone()).chain(pair.into_inner().flatten())).collect()
}

fn matched(source: &str) -> grameo::compat::Pairs<'_, Rule> {
    build_captures(source, true, |builder|
        super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder))
        .unwrap().into_typed_pairs()
}

fn direct_children<'a>(pair: &Pair<'a, Rule>) -> (Pair<'a, Rule>, Pair<'a, Rule>, Option<Pair<'a, Rule>>) {
    let mut children = pair.clone().into_inner();
    (children.next().unwrap(), children.next().unwrap(), children.next().map(|clause|
        clause.into_inner().next().unwrap()))
}

fn assert_anchor(actual: &Pair<'_, Rule>, expected: &Pair<'_, Rule>) {
    assert_eq!((actual.as_rule(), actual.as_span().start(), actual.as_span().end(), actual.as_str()),
        (expected.as_rule(), expected.as_span().start(), expected.as_span().end(), expected.as_str()));
}

#[test]
fn if_regions_keep_contextual_anchors_and_original_dangling_else_binding() {
    for source in [
        "int f(int flag){if(flag) return 1; return 0;}",
        "int f(int flag){if(flag) if(flag-1) return 2; else return 3; return 0;}",
        "int f(int flag){if(flag){if(flag-1) return 2;}else return 3; return 0;}",
        "int f(int flag){if((flag=read())!=0) {int local=flag; use(&local);} else while(flag--) use(flag); return flag;}",
        "int f(int flag){if(flag /*condition*/) ; else /*branch*/ {goto done;} done:return 0;}",
    ] {
        with_cache_scope(|| {
            let pairs = captures(matched(source));
            let ifs: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::if_statement).collect();
            assert!(!ifs.is_empty());
            for pair in ifs {
                let (condition, then_branch, else_branch) = direct_children(pair);
                for _ in 0..2 {
                    let value = take_if_region(pair).unwrap_or_else(|| panic!("missing if region: {}", pair.as_str()));
                    assert_anchor(&value.condition, &condition);
                    assert_anchor(&value.then_branch, &then_branch);
                    match (&value.else_branch, &else_branch) {
                        (Some(actual), Some(expected)) => assert_anchor(actual, expected),
                        (None, None) => {}, _ => panic!("changed else binding"),
                    }
                    // The receipt owns structure, not normalization. It must
                    // retain exactly the values supplied by ordered lowering.
                    let cond = Expression::new(ExprKind::Ident("normalized_condition".into()));
                    let then_body = vec![vybe_platform_libc::emitter::build::stmt(StmtKind::Return(None))];
                    let else_body = else_branch.as_ref().map(|_| Vec::new());
                    let expected = vybe_platform_libc::emitter::build::stmt(StmtKind::If {
                        cond: cond.clone(), then_body: then_body.clone(), elifs: Vec::new(), else_body: else_body.clone(),
                    });
                    assert_eq!(format!("{:?}", value.finish(cond, then_body, else_body)), format!("{expected:?}"));
                }
            }
            // Completed branch controls and typed expression syntax retain
            // their independent consumers even when an if receipt exists.
            for pair in pairs.iter().filter(|pair| pair.as_rule() == Rule::statement && pair.as_str() == "return 0;") {
                assert!(take_control_statement(pair).is_none());
            }
        });
    }
}

#[test]
fn if_regions_validate_all_anchors_before_spending_and_allow_reentrant_lowering() {
    let source = "int f(int x){if(x) break; else continue; if(x+1) return; return 0;}";
    with_cache_scope(|| {
        let pairs = captures(matched(source));
        let ifs: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::if_statement).collect();
        let pair = ifs[0];
        let input = (source.as_ptr() as usize, source.len());
        let span = pair.as_span();
        let key = (span.start(), span.end());
        let saved = CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            let foreign = values.ifs[&(ifs[1].as_span().start(), ifs[1].as_span().end())].recipe.condition;
            let entry = values.ifs.get_mut(&key).unwrap();
            let saved = entry.recipe;
            entry.recipe.else_branch = Some(foreign);
            saved
        });
        assert!(take_if_region(pair).is_none());
        CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let entry = cache.as_mut().unwrap().get_mut(&input).unwrap().ifs.get_mut(&key).unwrap();
            assert!(entry.first.is_some());
            entry.recipe = saved;
        });
        let value = take_if_region(pair).unwrap();
        assert!(matches!(take_control_statement(&value.then_branch), Some(OwnedStatement::Emit(_))));
        assert!(matches!(take_control_statement(value.else_branch.as_ref().unwrap()), Some(OwnedStatement::Emit(_))));
        with_cache_scope(|| { assert!(take_if_region(pair).is_none()); });
        assert!(take_if_region(pair).is_some());
    });
}

#[test]
fn if_regions_merge_only_committed_units_and_failed_reparse_invalidates() {
    let source = "int f(int x){if(x) return 1; return 0;} int g(int y){if(y) return 2; return 0;} int broken(){if(1)";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start| build_unit_captures(source, false, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "function_definition", source, start,
                MatchOptions::default(), builder));
        let first = captures(build(0).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::if_statement).unwrap();
        let second = captures(build(source.find("int g").unwrap()).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::if_statement).unwrap();
        assert!(build(source.find("int broken").unwrap()).is_err());
        assert!(take_if_region(&first).is_some());
        assert!(take_if_region(&second).is_some());
        assert!(build_captures(source, false, |builder|
            super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder)).is_err());
        assert!(take_if_region(&first).is_none());
        assert!(take_if_region(&second).is_none());
    });
}

#[test]
fn if_regions_preserve_capture_tree_and_budget_error_contract() {
    for source in ["int f(){if(1) return; else return;}", "int f(){if(1) return; else", "int f(){if(1"] {
        for work in [1, 100, usize::MAX] {
            let options = MatchOptions { max_steps: work, ..MatchOptions::default() };
            struct NoAst;
            impl Builder for NoAst {
                fn checkpoint(&self) -> usize { 0 }
                fn rollback(&mut self, _: usize) {}
                fn begin(&mut self, _: usize, _: usize) {}
                fn finish(&mut self, _: usize, _: Span, _: &str) {}
            }
            let mut plain = NoAst;
            let reference = super::super::build_hybrid_inner(Rule::program, source, options.clone(), &mut plain);
            with_cache_scope(|| {
                let actual = build_captures(source, false, |builder|
                    super::super::build_hybrid_inner(Rule::program, source, options, builder));
                match (reference, actual) {
                    (Ok(reference), Ok(actual)) => {
                        let shape = |pairs| captures(pairs).into_iter().map(|pair|
                            (pair.as_rule(), pair.as_span().start(), pair.as_span().end())).collect::<Vec<_>>();
                        assert_eq!(shape(reference.into_typed_pairs()), shape(actual.into_typed_pairs()));
                    }
                    (Err(reference), Err(actual)) => assert_eq!(format!("{reference:?}"), format!("{actual:?}")),
                    _ => panic!("changed grammar success/error contract"),
                }
            });
        }
    }
}
