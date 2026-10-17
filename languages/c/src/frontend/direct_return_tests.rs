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
fn return_regions_keep_exact_contextual_expression_and_replay_only_the_shell() {
    for source in [
        "int f(int x){return x;}",
        "int f(int x){return x++ + read(x);}",
        "int f(char *p,int i){return p[i++];}",
        "int *f(int *p){return &p[read()];}",
        "int f(int x){return x ? read() : other();}",
        "int f(int x){return (x++,x);}",
        "int f(int x){return sizeof(x++);}",
        "int f(int x){return ({int y=x; y+1;});}",
        "int f(){return ((int[]){3,5})[1];}",
        "void f(int x){if(x)return;return consume(x);}",
    ] {
        with_cache_scope(|| {
            let pairs = captures(matched(source));
            let mut valued = 0;
            for pair in pairs.iter().filter(|pair| pair.as_rule() == Rule::return_statement) {
                let expected = pair.clone().into_inner().next();
                let Some(expected) = expected else {
                    assert!(take_return_region(pair).is_none(), "void return keeps control handoff");
                    continue;
                };
                valued += 1;
                for _ in 0..2 {
                    let region = take_return_region(pair).unwrap();
                    let actual = &region.expression;
                    assert_eq!((actual.as_rule(), actual.as_span().start(), actual.as_span().end(), actual.as_str()),
                        (expected.as_rule(), expected.as_span().start(), expected.as_span().end(), expected.as_str()));
                    let normalized = Expression::new(ExprKind::Ident("contextual_return_value".into()));
                    let expected = vybe_platform_libc::emitter::build::stmt(StmtKind::Return(Some(normalized.clone())));
                    assert_eq!(format!("{:?}", region.finish(normalized)), format!("{expected:?}"));
                }
                assert!(take_return_region(&expected).is_none(), "only original return boundary is eligible");
            }
            assert!(valued > 0);
        });
    }
}

#[test]
fn return_shell_preserves_complete_contextually_normalized_ast() {
    for source in [
        "int f(char *p,int i){return p[i];}",
        "long long f(long long *p){return p[1];}",
        "unsigned short f(unsigned short *p){return *p;}",
        "int *f(int *p){return &p[2];}",
        "int f(int x){return (x++,x*2);}",
        "int f(int x){return x ? read(x) : other();}",
        "int f(int n){int a[n];return sizeof(a);}",
        "int f(int x){return ({int y=x;y+1;});}",
    ] {
        let expected = crate::parse_with_peg(source).unwrap();
        assert_eq!(format!("{:?}", crate::parse_optimized(source).unwrap()), format!("{expected:?}"));
        with_cache_scope(|| {
            let pairs = captures(matched(source));
            let pair = pairs.iter().find(|pair| pair.as_rule() == Rule::return_statement).unwrap();
            let mut actual = expected.clone();
            let body = actual.body.iter_mut().find_map(|statement| match &mut statement.kind {
                StmtKind::FunctionDecl { name, body, .. } if name == "f" => Some(body),
                _ => None,
            }).unwrap();
            let target = body.iter_mut().find(|statement| matches!(statement.kind, StmtKind::Return(Some(_)))).unwrap();
            let StmtKind::Return(Some(normalized)) = &target.kind else { unreachable!(); };
            let normalized = normalized.clone();
            for _ in 0..2 {
                let mut owned = take_return_region(pair).unwrap().finish(normalized.clone());
                // The common statement wrapper, not this producer, owns spans.
                owned.span = target.span.clone();
                assert_eq!(format!("{owned:?}"), format!("{target:?}"), "{source}");
                *target = owned;
            }
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}");
        });
    }
}

#[test]
fn return_regions_validate_before_take_and_restore_nested_source_scopes() {
    let source = "int f(int x){if(x)return first(x);return second(x);}";
    with_cache_scope(|| {
        let pairs = captures(matched(source));
        let returns: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::return_statement).collect();
        let input = (source.as_ptr() as usize, source.len());
        let key = (returns[0].as_span().start(), returns[0].as_span().end());
        let saved = CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            let foreign = values.returns[&(returns[1].as_span().start(), returns[1].as_span().end())].expression;
            std::mem::replace(&mut values.returns.get_mut(&key).unwrap().expression, foreign)
        });
        assert!(take_return_region(returns[0]).is_none());
        CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let entry = cache.as_mut().unwrap().get_mut(&input).unwrap().returns.get_mut(&key).unwrap();
            assert!(entry.first.is_some(), "failed anchor resolution must not consume owned AST");
            entry.expression = saved;
        });
        let outer = take_return_region(returns[0]).unwrap();
        with_cache_scope(|| {
            assert!(take_return_region(returns[0]).is_none());
            let nested = captures(matched(source));
            let pair = nested.iter().find(|pair| pair.as_rule() == Rule::return_statement).unwrap();
            assert!(take_return_region(pair).is_some());
        });
        assert!(take_return_region(returns[0]).is_some());
        assert!(take_return_region(returns[1]).is_some(), "no TLS borrow survives a take");
        let result = outer.finish(Expression::new(ExprKind::Lit(vybe_ast::Literal::Int(7))));
        assert!(matches!(result.kind, StmtKind::Return(Some(_))));
    });
}

#[test]
fn return_regions_append_committed_units_but_never_publish_failed_units() {
    let source = "int f(){return first();} int g(){return second();} int broken(){return (";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start| build_unit_captures(source, false, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "function_definition", source, start,
                MatchOptions::default(), builder));
        let first = captures(build(0).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::return_statement).unwrap();
        let second = captures(build(source.find("int g").unwrap()).unwrap().into_typed_pairs()).into_iter()
            .find(|pair| pair.as_rule() == Rule::return_statement).unwrap();
        assert!(build(source.find("int broken").unwrap()).is_err());
        assert!(take_return_region(&first).is_some());
        assert!(take_return_region(&second).is_some());
        let input = (source.as_ptr() as usize, source.len());
        assert_eq!(CACHE.with(|cache| cache.borrow().as_ref().unwrap()[&input].returns.len()), 2);
        assert!(build_captures(source, false, |builder|
            super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder)).is_err());
        assert!(take_return_region(&first).is_none());
        assert!(take_return_region(&second).is_none());
    });
}

#[test]
fn return_regions_preserve_capture_shapes_errors_and_budgets() {
    for source in [
        "int f(){return value();}", "void f(){return;}",
        "int f(){return ({int x=1;x;});}",
        "int f(){return (1+);}", "int f(){return value()}", "int f(){return",
    ] {
        for max_steps in [1, 100, usize::MAX] {
            let options = MatchOptions { max_steps, ..Default::default() };
            let control = engine::parse_program(&*GRAMMAR, "program", source, options.clone());
            let mut builder = Direct::new(source);
            let actual = engine::build_capture_program(&*GRAMMAR, "program", source, options, &mut builder);
            match (control, actual) {
                (Ok(control), Ok(actual)) => {
                    let shape = |pairs| captures(pairs).into_iter().map(|pair|
                        (pair.as_rule(), pair.as_span().start(), pair.as_span().end(), pair.as_str().to_string())).collect::<Vec<_>>();
                    assert_eq!(shape(control.into_typed_pairs()), shape(actual.into_typed_pairs()));
                }
                (Err(expected), Err(actual)) => {
                    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                    assert!(builder.into_initializers().returns.is_empty());
                }
                _ => panic!("return hook changed grammar acceptance: {source}"),
            }
        }
    }
}
