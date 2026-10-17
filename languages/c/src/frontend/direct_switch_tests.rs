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

fn assert_anchor(actual: &Pair<'_, Rule>, expected: &Pair<'_, Rule>) {
    assert_eq!((actual.as_rule(), actual.as_span().start(), actual.as_span().end(), actual.as_str()),
        (expected.as_rule(), expected.as_span().start(), expected.as_span().end(), expected.as_str()));
}

#[test]
fn switch_regions_keep_selector_and_body_anchors_without_collecting_or_promoting_cases() {
    for source in [
        "int f(int x){switch(x){case 0:return 1; default:return 2;}}",
        "int f(int x){switch(x++) {default:use(x); case 1:case 2:use(3);break;}return x;}",
        "int f(unsigned long long x){switch(x){case 4294967296ULL:break;case 2 ... 4:break;}return 0;}",
        "int f(int x){switch(read()) case 0: return x;return 0;}",
        "int f(int x){switch(x){case 1: switch(x-1){case 0:break;}break;default:while(x--) break;}return 0;}",
        "int f(int x){switch(x /*selector*/) {case 0:{int y=x;use(&y);}default:break;}return x;}",
    ] {
        with_cache_scope(|| {
            let pairs = captures(matched(source));
            let switches: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::switch_statement).collect();
            assert!(!switches.is_empty());
            for pair in switches {
                let expected: Vec<_> = pair.clone().into_inner().collect();
                assert_eq!(expected.len(), 2);
                for _ in 0..2 {
                    let region = take_switch_region(pair).unwrap();
                    assert_anchor(&region.subject, &expected[0]);
                    assert_anchor(&region.body, &expected[1]);
                    let expr = Expression::new(ExprKind::Ident("promoted_selector".into()));
                    let cases = vec![SwitchCase {
                        conditions: vec![vybe_ast::CaseCondition::Value(Expression::new(ExprKind::Ident("promoted_case".into())))],
                        body: vec![vybe_platform_libc::emitter::build::stmt(StmtKind::Break(BreakTarget::Implicit))],
                    }];
                    let default = Some(vec![vybe_platform_libc::emitter::build::stmt(StmtKind::Return(None))]);
                    let expected = vybe_platform_libc::emitter::build::stmt(StmtKind::Switch {
                        expr: expr.clone(), cases: cases.clone(), default: default.clone(),
                    });
                    assert_eq!(format!("{:?}", region.finish(expr, cases, default)), format!("{expected:?}"));
                }
            }
            for pair in pairs.iter().filter(|pair| pair.as_rule() == Rule::while_statement) {
                assert!(take_while_region(pair).is_some());
            }
        });
    }
}

#[test]
fn switch_regions_validate_body_before_take_and_allow_reentrant_nested_switches() {
    let source = "int f(int x){switch(x){case 0:break;}switch(x+1){default:break;}return 0;}";
    with_cache_scope(|| {
        let pairs = captures(matched(source));
        let switches: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::switch_statement).collect();
        let input = (source.as_ptr() as usize, source.len());
        let key = (switches[0].as_span().start(), switches[0].as_span().end());
        let saved = CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            let foreign = values.switches[&(switches[1].as_span().start(), switches[1].as_span().end())].recipe.body;
            let entry = values.switches.get_mut(&key).unwrap();
            let saved = entry.recipe;
            entry.recipe.body = foreign;
            saved
        });
        assert!(take_switch_region(switches[0]).is_none());
        CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let entry = cache.as_mut().unwrap().get_mut(&input).unwrap().switches.get_mut(&key).unwrap();
            assert!(entry.first.is_some());
            entry.recipe = saved;
        });
        let region = take_switch_region(switches[0]).unwrap();
        assert!(take_switch_region(switches[1]).is_some());
        assert!(take_switch_region(&region.body).is_none());
        with_cache_scope(|| { assert!(take_switch_region(switches[0]).is_none()); });
        assert!(take_switch_region(switches[0]).is_some());
    });
}

#[test]
fn switch_regions_append_only_committed_units_and_invalidate_failed_full_reparse() {
    let source = "int f(int x){switch(x){case 0:break;}} int g(int y){switch(y){default:break;}} int broken(){switch(";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start| build_unit_captures(source, false, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "function_definition", source, start,
                MatchOptions::default(), builder));
        let first = captures(build(0).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::switch_statement).unwrap();
        let second = captures(build(source.find("int g").unwrap()).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::switch_statement).unwrap();
        assert!(build(source.find("int broken").unwrap()).is_err());
        assert!(take_switch_region(&first).is_some());
        assert!(take_switch_region(&second).is_some());
        assert!(build_captures(source, false, |builder|
            super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder)).is_err());
        assert!(take_switch_region(&first).is_none());
        assert!(take_switch_region(&second).is_none());
    });
}

#[test]
fn switch_regions_preserve_capture_shape_and_bounded_error_contract() {
    for source in ["int f(int x){switch(x){case 0:break;default:break;}}", "int f(){switch(1)", "int f(){switch("] {
        for max_steps in [1, 100, usize::MAX] {
            let options = MatchOptions { max_steps, ..Default::default() };
            let control = engine::parse_program(&*GRAMMAR, "program", source, options.clone());
            let mut builder = Direct::new(source);
            let actual = engine::build_capture_program(&*GRAMMAR, "program", source, options, &mut builder);
            match (control, actual) {
                (Ok(control), Ok(actual)) => {
                    let shape = |pairs| captures(pairs).into_iter().map(|pair|
                        (pair.as_rule(), pair.as_span().start(), pair.as_span().end())).collect::<Vec<_>>();
                    assert_eq!(shape(control.into_typed_pairs()), shape(actual.into_typed_pairs()));
                    assert_eq!(builder.into_initializers().switches.len(), 1);
                }
                (Err(expected), Err(actual)) => {
                    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                    assert!(builder.into_initializers().switches.is_empty());
                }
                _ => panic!("switch hook changed acceptance: {source}"),
            }
        }
    }
}
