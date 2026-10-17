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
fn for_regions_preserve_ordered_clause_wrappers_including_empty_clauses() {
    for source in [
        "int f(){for(;;) break; return 0;}",
        "int f(){for(int i=0;i<3;++i) use(i); return 0;}",
        "int f(int i){for(i=read(); check(i); i=next(i)) {int j=i; use(&j);} return i;}",
        "int f(int i){for(/*init*/;/*condition*/;/*update*/) {if(i) continue; break;} return i;}",
        "int f(int i){for(;i;) for(int j=0;j<i;j++) do {use(j);} while(0); return i;}",
        "int f(){for(int i=0,j=2; i<j; ++i,--j) ; return 0;}",
    ] {
        with_cache_scope(|| {
            let pairs = captures(matched(source));
            let fors: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::for_statement).collect();
            assert!(!fors.is_empty());
            for pair in fors {
                let expected: Vec<_> = pair.clone().into_inner().collect();
                for _ in 0..2 {
                    let region = take_for_region(pair).unwrap_or_else(|| panic!("missing for: {}", pair.as_str()));
                    let mut actual = vec![region.init.clone(), region.condition.clone()];
                    actual.extend(region.update.clone());
                    actual.push(region.body.clone());
                    assert_eq!(actual.len(), expected.len());
                    for (actual, expected) in actual.iter().zip(&expected) { assert_anchor(actual, expected); }
                    // C type/scoping policy is supplied by the caller, not by
                    // the matcher. Preserve the actual normalized AST verbatim.
                    let init = Some(Box::new(vybe_platform_libc::emitter::build::stmt(StmtKind::Return(None))));
                    let cond = Some(Expression::new(ExprKind::Ident("normalized_condition".into())));
                    let update = Some(Expression::new(ExprKind::Ident("normalized_update".into())));
                    let body = vec![vybe_platform_libc::emitter::build::stmt(StmtKind::Break(BreakTarget::Implicit))];
                    let expected = vybe_platform_libc::emitter::build::stmt(StmtKind::For {
                        init: init.clone(), cond: cond.clone(), update: update.clone(), body: body.clone(),
                    });
                    assert_eq!(format!("{:?}", region.finish(init, cond, update, body)), format!("{expected:?}"));
                    assert!(matches!(take_for_region(pair).unwrap().finish(None, None, None, Vec::new()).kind,
                        StmtKind::For { init: None, cond: None, update: None, body } if body.is_empty()));
                }
            }
            for pair in pairs.iter().filter(|pair| pair.as_rule() == Rule::if_statement) {
                assert!(take_if_region(pair).is_some());
            }
            for pair in pairs.iter().filter(|pair| pair.as_rule() == Rule::do_while_statement) {
                assert!(take_do_while_region(pair).is_some());
            }
        });
    }
}

#[test]
fn for_regions_validate_all_clauses_before_take_and_preserve_reentrant_body_consumers() {
    let source = "int f(int x){for(;x;--x) break; for(;;) continue; return 0;}";
    with_cache_scope(|| {
        let pairs = captures(matched(source));
        let fors: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::for_statement).collect();
        let input = (source.as_ptr() as usize, source.len());
        let key = (fors[0].as_span().start(), fors[0].as_span().end());
        let saved = CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            let foreign = values.fors[&(fors[1].as_span().start(), fors[1].as_span().end())].recipe.body;
            let entry = values.fors.get_mut(&key).unwrap();
            let saved = entry.recipe;
            entry.recipe.update = Some(foreign);
            saved
        });
        assert!(take_for_region(fors[0]).is_none());
        CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let entry = cache.as_mut().unwrap().get_mut(&input).unwrap().fors.get_mut(&key).unwrap();
            assert!(entry.first.is_some());
            entry.recipe = saved;
        });
        let region = take_for_region(fors[0]).unwrap();
        assert!(matches!(take_control_statement(&region.body), Some(OwnedStatement::Emit(_))));
        with_cache_scope(|| { assert!(take_for_region(fors[0]).is_none()); });
        assert!(take_for_region(fors[0]).is_some());
        assert!(take_for_region(&region.init).is_none());
    });
}

#[test]
fn for_regions_append_committed_units_and_failed_reparse_invalidates_receipts() {
    let source = "int f(){for(;;) break;} int g(){for(int y=0;y<2;++y) continue;} int broken(){for(";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start| build_unit_captures(source, false, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "function_definition", source, start,
                MatchOptions::default(), builder));
        let first = captures(build(0).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::for_statement).unwrap();
        let second = captures(build(source.find("int g").unwrap()).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::for_statement).unwrap();
        assert!(build(source.find("int broken").unwrap()).is_err());
        assert!(take_for_region(&first).is_some());
        assert!(take_for_region(&second).is_some());
        assert!(build_captures(source, false, |builder|
            super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder)).is_err());
        assert!(take_for_region(&first).is_none());
        assert!(take_for_region(&second).is_none());
    });
}

#[test]
fn for_regions_keep_capture_shape_and_bounded_error_contract() {
    for source in ["int f(){for(int i=0;i<2;++i) continue;}", "int f(){for(int i=0;", "int f(){for(;;)"] {
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
                    assert_eq!(builder.into_initializers().fors.len(), 1);
                }
                (Err(expected), Err(actual)) => {
                    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                    assert!(builder.into_initializers().fors.is_empty());
                }
                _ => panic!("for hook changed acceptance: {source}"),
            }
        }
    }
}
