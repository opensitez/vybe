use super::*;
use grameo::{CompileOptions, CompiledGrammar, MatchOptions, engine};
use std::sync::LazyLock;

static GRAMMAR: LazyLock<CompiledGrammar> = LazyLock::new(||
    grameo::compile_with_options(include_str!("../grammar.pest"), CompileOptions::default()).unwrap());

fn captures<'a>(pairs: grameo::compat::Pairs<'a, Rule>) -> Vec<Pair<'a, Rule>> {
    pairs.flat_map(|pair| std::iter::once(pair.clone()).chain(pair.into_inner().flatten())).collect()
}

#[test]
fn function_regions_preserve_all_original_parts_and_shared_constructor_flags() {
    for source in ["int f(int x){return x;}", "typedef int T;T f(T x){return x;}",
        "int *f(int *p){return p;}", "int f(x) int x; {return x;}",
        "int outer(void){int inner(int x){return x;}return inner(3);}"]
    {
        with_cache_scope(|| {
            let pairs = build_captures(source, true, |builder|
                super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder))
                .unwrap().into_typed_pairs();
            for pair in captures(pairs).iter().filter(|pair| pair.as_rule() == Rule::function_definition) {
                let expected: Vec<_> = pair.clone().into_inner().collect();
                for _ in 0..2 {
                    let region = take_function_region(pair).unwrap_or_else(|| panic!("missing function: {}", pair.as_str()));
                    let parts: Vec<_> = region.parts().collect();
                    assert_eq!(parts.len(), expected.len());
                    for (part, expected) in parts.iter().zip(&expected) {
                        assert_eq!((part.as_rule(), part.as_span().start(), part.as_span().end(), part.as_str()),
                            (expected.as_rule(), expected.as_span().start(), expected.as_span().end(), expected.as_str()));
                    }
                    let params = vec![crate::walker::construct_plain_parameter("normalized".into())];
                    let body = vec![vybe_platform_libc::emitter::build::stmt(StmtKind::Return(None))];
                    let value = region.finish("emitted".into(), params.clone(), Some("int".into()), body.clone());
                    let expected = crate::walker::construct_function_declaration("emitted".into(), params, Some("int".into()), body);
                    assert_eq!(format!("{value:?}"), format!("{expected:?}"));
                    assert!(take_body(parts.last().unwrap()).is_some(), "body ownership must survive function shell");
                }
            }
        });
    }
}

#[test]
fn function_regions_validate_before_take_and_preserve_committed_units_and_reentrant_scopes() {
    let source = "int f(int first){return first;} int g(int second){return second;} int broken(){";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start| build_unit_captures(source, false, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "function_definition", source, start,
                MatchOptions::default(), builder));
        let first = build(0).unwrap().into_typed_pairs::<Rule>().next().unwrap();
        let second = build(source.find("int g").unwrap()).unwrap().into_typed_pairs::<Rule>().next().unwrap();
        assert!(build(source.find("int broken").unwrap()).is_err());
        let input = (source.as_ptr() as usize, source.len());
        let key = (first.as_span().start(), first.as_span().end());
        let saved = CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            let foreign = *values.functions[&(second.as_span().start(), second.as_span().end())].parts.last().unwrap();
            let entry = values.functions.get_mut(&key).unwrap();
            let last = entry.parts.last_mut().unwrap();
            std::mem::replace(last, foreign)
        });
        assert!(take_function_region(&first).is_none());
        CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let entry = cache.as_mut().unwrap().get_mut(&input).unwrap().functions.get_mut(&key).unwrap();
            assert!(entry.first.is_some());
            *entry.parts.last_mut().unwrap() = saved;
        });
        assert!(take_function_region(&first).is_some());
        assert!(take_function_region(&second).is_some());
        with_cache_scope(|| { assert!(take_function_region(&first).is_none()); });
        assert!(take_function_region(&first).is_some());
        assert!(build_captures(source, true, |builder|
            super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder)).is_err());
        assert!(take_function_region(&first).is_none());
        assert!(take_function_region(&second).is_none());
    });
}

#[test]
fn function_regions_keep_capture_shapes_and_bounded_grammar_failures() {
    for source in ["int f(int x){return x;}", "int f(int x){return x;} @", "int f(int x){"] {
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
                    assert_eq!(builder.into_initializers().functions.len(), 1);
                }
                (Err(expected), Err(actual)) => {
                    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                    assert!(builder.into_initializers().functions.is_empty());
                }
                _ => panic!("function hook changed grammar acceptance: {source}"),
            }
        }
    }
}
