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

#[test]
fn call_regions_preserve_exact_ordered_assignment_anchors_without_argument_policy() {
    for source in [
        "int f(){return target();}",
        "int f(int x){return target(x, x+1, x=2, (x,3));}",
        "int f(int x){return outer(inner(x++), x?left():right(), &x);}",
        "int f(int (*fp)(int), int x){return fp(x);}",
        "struct S { int (*run)(int); }; int f(struct S *p,int x){return p->run(x);}",
        "int f(char *p){return sscanf(p, \"%d\", &p[0]);}",
        "int f(int x){return target(/*first*/x, /*second*/ x+1 /*tail*/);}",
        "int f(int x){return table[x++](x)[x];}",
    ] {
        with_cache_scope(|| {
            let pairs = captures(matched(source));
            let calls: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::call_suffix).collect();
            assert!(!calls.is_empty());
            for pair in calls {
                let expected: Vec<_> = pair.clone().into_inner().flat_map(|list| list.into_inner()).collect();
                for _ in 0..2 {
                    let region = take_call_region(pair).unwrap_or_else(|| panic!("missing call: {}", pair.as_str()));
                    assert_eq!(region.arguments.len(), expected.len());
                    for (actual, expected) in region.arguments.iter().zip(&expected) {
                        assert_eq!(actual.as_rule(), Rule::assignment_expression);
                        assert_eq!((actual.as_span().start(), actual.as_span().end(), actual.as_str()),
                            (expected.as_span().start(), expected.as_span().end(), expected.as_str()));
                    }
                    let callee = Expression::new(ExprKind::Ident("normalized_callee".into()));
                    let args: Vec<_> = expected.iter().enumerate().map(|(index, _)|
                        Argument::positional(Expression::new(ExprKind::Ident(format!("normalized_arg{index}"))))).collect();
                    let expected = Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()), args: args.clone(), optional: false,
                    });
                    assert_eq!(format!("{:?}", region.finish(callee, args)), format!("{expected:?}"));
                }
            }
        });
    }
}

#[test]
fn call_regions_validate_all_arguments_before_take_and_preserve_reentrant_nested_consumers() {
    let source = "int f(int x){return outer(inner(x),x+1)+other(x);}";
    with_cache_scope(|| {
        let pairs = captures(matched(source));
        let outer = pairs.iter().find(|pair| pair.as_rule() == Rule::call_suffix && pair.as_str() == "(inner(x),x+1)").unwrap();
        let nested = pairs.iter().find(|pair| pair.as_rule() == Rule::call_suffix && pair.as_str() == "(x)").unwrap();
        let other = pairs.iter().rev().find(|pair| pair.as_rule() == Rule::call_suffix).unwrap();
        let input = (source.as_ptr() as usize, source.len());
        let key = (outer.as_span().start(), outer.as_span().end());
        let saved = CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            let foreign = values.calls[&(other.as_span().start(), other.as_span().end())].arguments[0];
            let entry = values.calls.get_mut(&key).unwrap();
            std::mem::replace(&mut entry.arguments[1], foreign)
        });
        assert!(take_call_region(outer).is_none());
        CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let entry = cache.as_mut().unwrap().get_mut(&input).unwrap().calls.get_mut(&key).unwrap();
            assert!(entry.first.is_some());
            entry.arguments[1] = saved;
        });
        let region = take_call_region(outer).unwrap();
        assert!(take_call_region(nested).is_some());
        assert!(take_call_region(&region.arguments[0]).is_none());
        with_cache_scope(|| { assert!(take_call_region(outer).is_none()); });
        assert!(take_call_region(outer).is_some());
    });
}

#[test]
fn call_regions_append_only_committed_units_and_failed_reparse_invalidates() {
    let source = "int f(){return first();} int g(int x){return second(x);} int broken(){return third(";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start| build_unit_captures(source, false, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "function_definition", source, start,
                MatchOptions::default(), builder));
        let first = captures(build(0).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::call_suffix).unwrap();
        let second = captures(build(source.find("int g").unwrap()).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::call_suffix).unwrap();
        assert!(build(source.find("int broken").unwrap()).is_err());
        assert!(take_call_region(&first).is_some());
        assert!(take_call_region(&second).is_some());
        assert!(build_captures(source, false, |builder|
            super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder)).is_err());
        assert!(take_call_region(&first).is_none());
        assert!(take_call_region(&second).is_none());
    });
}

#[test]
fn call_regions_preserve_capture_shape_and_bounded_failure_contract() {
    for source in ["int f(int x){return outer(inner(x),x+1);}", "int f(){return call(1,", "int f(){return call("] {
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
                    assert_eq!(builder.into_initializers().calls.len(), 2);
                }
                (Err(expected), Err(actual)) => {
                    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                    assert!(builder.into_initializers().calls.is_empty());
                }
                _ => panic!("call hook changed acceptance: {source}"),
            }
        }
    }
}
