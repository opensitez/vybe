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
fn loop_regions_keep_exact_condition_body_anchors_and_ordinary_constructor_flags() {
    for source in [
        "int f(int x){while(x) x--; do x++; while(x<2); return x;}",
        "int f(int x){while(x=read()) {int y=x; use(&y);} do {use(x++);} while(read()); return x;}",
        "int f(int x){while(x) do {if(x) continue; else break;} while(--x); return x;}",
        "int f(){while(1); do {continue;} while(0); return 0;}",
        "int f(int x){do /*body*/ ; while(x /*condition*/); while(x) /*body*/ ; return x;}",
    ] {
        with_cache_scope(|| {
            let pairs = captures(matched(source));
            let loops: Vec<_> = pairs.iter().filter(|pair|
                matches!(pair.as_rule(), Rule::while_statement | Rule::do_while_statement)).collect();
            assert!(!loops.is_empty());
            for pair in loops {
                let do_while = pair.as_rule() == Rule::do_while_statement;
                let mut children = pair.clone().into_inner();
                let first = children.next().unwrap();
                let second = children.next().unwrap();
                assert!(children.next().is_none());
                let (condition, body_anchor) = if do_while { (second, first) } else { (first, second) };
                for _ in 0..2 {
                    let cond = Expression::new(ExprKind::Ident("already_normalized".into()));
                    let body = vec![vybe_platform_libc::emitter::build::stmt(StmtKind::Return(None))];
                    let (value, kind) = if do_while {
                        assert!(take_while_region(pair).is_none());
                        let region = take_do_while_region(pair).unwrap();
                        assert_anchor(&region.condition, &condition);
                        assert_anchor(&region.body, &body_anchor);
                        (region.finish(cond.clone(), body.clone()), StmtKind::DoWhile { body, cond, until: false })
                    } else {
                        assert!(take_do_while_region(pair).is_none());
                        let region = take_while_region(pair).unwrap();
                        assert_anchor(&region.condition, &condition);
                        assert_anchor(&region.body, &body_anchor);
                        (region.finish(cond.clone(), body.clone()), StmtKind::While { cond, body, else_body: None })
                    };
                    assert_eq!(format!("{value:?}"), format!("{:?}", vybe_platform_libc::emitter::build::stmt(kind)));
                }
            }
            // Hook observation must not consume independent if/body controls.
            for pair in pairs.iter().filter(|pair| pair.as_rule() == Rule::if_statement) {
                assert!(take_if_region(pair).is_some());
            }
        });
    }
}

#[test]
fn loop_regions_validate_late_anchor_before_take_and_keep_reentrant_body_handoffs() {
    let source = "int f(int x){while(x) break; do continue; while(x+1); return 0;}";
    with_cache_scope(|| {
        let pairs = captures(matched(source));
        let while_pair = pairs.iter().find(|pair| pair.as_rule() == Rule::while_statement).unwrap();
        let do_pair = pairs.iter().find(|pair| pair.as_rule() == Rule::do_while_statement).unwrap();
        let input = (source.as_ptr() as usize, source.len());
        let key = (Rule::while_statement, while_pair.as_span().start(), while_pair.as_span().end());
        let saved = CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            let foreign = values.loops[&(Rule::do_while_statement, do_pair.as_span().start(), do_pair.as_span().end())].recipe.body;
            let entry = values.loops.get_mut(&key).unwrap();
            let saved = entry.recipe;
            entry.recipe.body = foreign;
            saved
        });
        assert!(take_while_region(while_pair).is_none());
        CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let entry = cache.as_mut().unwrap().get_mut(&input).unwrap().loops.get_mut(&key).unwrap();
            assert!(entry.first.is_some());
            entry.recipe = saved;
        });
        let region = take_while_region(while_pair).unwrap();
        assert!(matches!(take_control_statement(&region.body), Some(OwnedStatement::Emit(_))));
        let region = take_do_while_region(do_pair).unwrap();
        assert!(matches!(take_control_statement(&region.body), Some(OwnedStatement::Emit(_))));
        with_cache_scope(|| { assert!(take_while_region(while_pair).is_none()); });
        assert!(take_while_region(while_pair).is_some());
    });
}

#[test]
fn loop_regions_append_committed_units_and_failed_reparse_invalidates_both_kinds() {
    let source = "int f(int x){while(x) break;} int g(int y){do continue; while(y);} int broken(){do {";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start| build_unit_captures(source, false, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "function_definition", source, start,
                MatchOptions::default(), builder));
        let first = captures(build(0).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::while_statement).unwrap();
        let second = captures(build(source.find("int g").unwrap()).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::do_while_statement).unwrap();
        assert!(build(source.find("int broken").unwrap()).is_err());
        assert!(take_while_region(&first).is_some());
        assert!(take_do_while_region(&second).is_some());
        assert!(build_captures(source, false, |builder|
            super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder)).is_err());
        assert!(take_while_region(&first).is_none());
        assert!(take_do_while_region(&second).is_none());
    });
}

#[test]
fn loop_regions_preserve_capture_shape_and_bounded_failure_contract() {
    for source in ["int f(){while(1) break; do continue; while(0);}", "int f(){while(1)", "int f(){do;while("] {
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
                    assert_eq!(builder.into_initializers().loops.len(), 2);
                }
                (Err(expected), Err(actual)) => {
                    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                    assert!(builder.into_initializers().loops.is_empty());
                }
                _ => panic!("loop hook changed acceptance: {source}"),
            }
        }
    }
}
