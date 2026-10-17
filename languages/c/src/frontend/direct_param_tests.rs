use super::*;
use grameo::{CompileOptions, CompiledGrammar, MatchOptions, engine};
use std::sync::LazyLock;

static GRAMMAR: LazyLock<CompiledGrammar> = LazyLock::new(||
    grameo::compile_with_options(include_str!("../grammar.pest"), CompileOptions::default()).unwrap());

fn captures<'a>(pairs: grameo::compat::Pairs<'a, Rule>) -> Vec<Pair<'a, Rule>> {
    pairs.flat_map(|pair| std::iter::once(pair.clone()).chain(pair.into_inner().flatten())).collect()
}

fn match_parameters(source: &str) -> grameo::compat::Pairs<'_, Rule> {
    build_captures(source, false, |builder|
        super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder))
        .unwrap().into_typed_pairs()
}

#[test]
fn plain_parameters_own_shared_ordinary_ast_and_exact_specifier_anchors() {
    let source = "typedef unsigned short Word; int f(const int value, Word count, double real, int /*gap*/ name /*tail*/);";
    with_cache_scope(|| {
        let pairs = captures(match_parameters(source));
        let parameters: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::parameter_decl).collect();
        assert_eq!(parameters.len(), 4);
        for ((pair, name), specs) in parameters.into_iter().zip(["value", "count", "real", "name"])
            .zip(["const int", "Word", "double", "int /*gap*/"])
        {
            let value = take_plain_parameter(pair).unwrap_or_else(|| panic!("missing bare parameter: {}", pair.as_str()));
            assert_eq!(format!("{:?}", value.parameter), format!("{:?}", crate::walker::construct_plain_parameter(name.into())));
            assert_eq!(value.specifiers.as_rule(), Rule::declaration_specifiers);
            assert_eq!(value.specifiers.as_str().trim(), specs);
            let reference_specifiers = pair.clone().into_inner()
                .find(|child| child.as_rule() == Rule::declaration_specifiers).unwrap();
            assert_eq!((value.specifiers.as_span().start(), value.specifiers.as_span().end(),
                value.specifiers.as_str()), (reference_specifiers.as_span().start(),
                reference_specifiers.as_span().end(), reference_specifiers.as_str()));
            assert_eq!(value.declaration_span, Span { start: pair.as_span().start(), end: pair.as_span().end() });
            let declarator = pair.clone().into_inner().find(|child| child.as_rule() == Rule::declarator).unwrap();
            assert_eq!(value.declarator_span, Span { start: declarator.as_span().start(), end: declarator.as_span().end() });
            assert!(value.parameter.type_hint.is_none());
            assert!(value.parameter.default.is_none());
            let replay = take_plain_parameter(pair).unwrap();
            assert_eq!(format!("{:?}", replay.parameter), format!("{:?}", value.parameter));
        }
    });
}

#[test]
fn plain_parameters_decline_explicit_declarator_shapes_and_complex_specifiers() {
    for declaration in ["int *p", "int **p", "int *const p", "int (p)", "int (*p)[3]",
        "int p[3]", "int p[]", "int (*callback)(int)", "int p(int)",
        "int", "int *", "int [3]", "void", "struct S value", "enum E value", "typeof(1) value",
        "__attribute__((unused)) int value", "int __attribute__((unused)) value",
        "int __attribute__((unused, cold)) value", "_Alignas(8) int value",
        "struct __attribute__((packed)) S value"]
    {
        let source = format!("int f({declaration});");
        with_cache_scope(|| {
            let pairs = captures(match_parameters(&source));
            let pair = pairs.iter().find(|pair| pair.as_rule() == Rule::parameter_decl && pair.as_str() == declaration)
                .unwrap_or_else(|| panic!("missing contextual parameter control: {declaration}"));
            assert!(take_plain_parameter(pair).is_none(), "unsupported syntax certified ordinary: {declaration}");
        });
    }
    // Alias classification is deliberately not a matcher decision. The ordered
    // consumer must reject these syntax receipts under its resolved-type guard.
    let source = "typedef int *P; typedef int A[3]; typedef int (*F)(int); int f(P p,A a,F callback);";
    with_cache_scope(|| {
        let pairs = captures(match_parameters(source));
        for name in ["P p", "A a", "F callback"] {
            let pair = pairs.iter().find(|pair| pair.as_rule() == Rule::parameter_decl && pair.as_str() == name).unwrap();
            assert!(take_plain_parameter(pair).unwrap().parameter.type_hint.is_none());
        }
    });
}

#[test]
fn plain_parameters_validate_before_take_and_merge_only_committed_units() {
    let source = "int f(int first){return first;} int g(int second){return second;} int broken(int partial){";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start| build_unit_captures(source, false, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "function_definition", source, start,
                MatchOptions::default(), builder));
        let first = captures(build(0).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::parameter_decl).unwrap();
        let second = captures(build(source.find("int g").unwrap()).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::parameter_decl).unwrap();
        assert!(build(source.find("int broken").unwrap()).is_err());
        let key = (first.as_span().start(), first.as_span().end());
        let saved = CACHE.with(|slot| {
            let mut slot = slot.borrow_mut();
            let cache = slot.as_mut().unwrap().get_mut(&(source.as_ptr() as usize, source.len())).unwrap();
            let foreign = cache.parameters.get(&(second.as_span().start(), second.as_span().end())).unwrap().recipe.specifiers;
            std::mem::replace(&mut cache.parameters.get_mut(&key).unwrap().recipe.specifiers, foreign)
        });
        assert!(take_plain_parameter(&first).is_none());
        CACHE.with(|slot| {
            let mut slot = slot.borrow_mut();
            let cache = slot.as_mut().unwrap().get_mut(&(source.as_ptr() as usize, source.len())).unwrap();
            let entry = cache.parameters.get_mut(&key).unwrap();
            assert!(entry.first.is_some(), "foreign anchor spent the owned parameter");
            entry.recipe.specifiers = saved;
        });
        assert_eq!(take_plain_parameter(&first).unwrap().parameter.name, "first");
        assert_eq!(take_plain_parameter(&second).unwrap().parameter.name, "second");
        with_cache_scope(|| assert!(take_plain_parameter(&first).is_none()));
        assert_eq!(take_plain_parameter(&first).unwrap().parameter.name, "first");
        assert!(build_captures(source, false, |builder|
            engine::build_capture_program(&*GRAMMAR, "program", source, MatchOptions::default(), builder)).is_err());
        assert!(take_plain_parameter(&first).is_none());
    });
}

#[test]
fn plain_parameters_rollback_and_budget_failures_publish_no_partial_parameters() {
    for source in ["int f(int first){return first;}", "int f(int first){return first;} @", "int f(int partial){"] {
        for max_steps in [1, 100, usize::MAX] {
            let options = MatchOptions { max_steps, ..Default::default() };
            let control = engine::parse_program(&*GRAMMAR, "program", source, options.clone());
            let mut builder = Direct::new(source);
            let actual = engine::build_capture_program(&*GRAMMAR, "program", source, options, &mut builder);
            match (control, actual) {
                (Ok(_), Ok(_)) => assert_eq!(builder.into_initializers().parameters.len(), 1),
                (Err(expected), Err(actual)) => {
                    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                    assert!(builder.into_initializers().parameters.is_empty());
                }
                _ => panic!("parameter hook changed acceptance: {source}"),
            }
        }
    }
}
