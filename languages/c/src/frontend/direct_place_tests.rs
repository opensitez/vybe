use super::*;
use grameo::{CompileOptions, CompiledGrammar, MatchOptions, engine};
use std::sync::LazyLock;

static GRAMMAR: LazyLock<CompiledGrammar> = LazyLock::new(||
    grameo::compile_with_options(include_str!("../grammar.pest"), CompileOptions::default()).unwrap());

fn captures<'a>(pairs: grameo::compat::Pairs<'a, Rule>) -> Vec<Pair<'a, Rule>> {
    pairs.flat_map(|pair| std::iter::once(pair.clone()).chain(pair.into_inner().flatten())).collect()
}

#[test]
fn place_regions_validate_before_take_and_invalidate_failed_reparse() {
    let source = "int f(int i){return a[i].x+b[i+1].y;}";
    with_cache_scope(|| {
        let pairs = build_captures(source, false, |builder|
            super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder))
            .unwrap().into_typed_pairs();
        let pairs = captures(pairs);
        let indices: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::index_suffix).collect();
        let members: Vec<_> = pairs.iter().filter(|pair| pair.as_rule() == Rule::member_suffix).collect();
        let input = (source.as_ptr() as usize, source.len());
        let ikey = (indices[0].as_span().start(), indices[0].as_span().end());
        let mkey = (Rule::member_suffix, members[0].as_span().start(), members[0].as_span().end());
        let saved = CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            let foreign = values.indices[&(indices[1].as_span().start(), indices[1].as_span().end())].anchor;
            let entry = values.indices.get_mut(&ikey).unwrap();
            std::mem::replace(&mut entry.anchor, foreign)
        });
        assert!(take_index_region(indices[0]).is_none());
        CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            assert!(values.indices[&ikey].first.is_some());
            values.indices.get_mut(&ikey).unwrap().anchor = saved;
        });
        assert!(take_index_region(indices[0]).is_some());
        let saved = CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            let foreign = values.members[&(Rule::member_suffix, members[1].as_span().start(), members[1].as_span().end())].anchor;
            std::mem::replace(&mut values.members.get_mut(&mkey).unwrap().anchor, foreign)
        });
        assert!(take_member_region(members[0]).is_none());
        CACHE.with(|cache| {
            let mut cache = cache.borrow_mut();
            let values = cache.as_mut().unwrap().get_mut(&input).unwrap();
            assert!(values.members[&mkey].first.is_some());
            values.members.get_mut(&mkey).unwrap().anchor = saved;
        });
        assert!(take_member_region(members[0]).is_some());
        with_cache_scope(|| { assert!(take_index_region(indices[0]).is_none()); });
        assert!(take_index_region(indices[0]).is_some());
        let failure: Result<(), ()> = build_captures(source, false, |_| Err(()));
        assert!(failure.is_err());
        assert!(take_index_region(indices[0]).is_none());
        assert!(take_member_region(members[0]).is_none());
    });
}

#[test]
fn place_regions_preserve_unit_transactions_capture_shapes_and_bounded_errors() {
    let source = "int f(){return a[0].x;} int g(){return b[1].y;} int broken(){return c[";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start| build_unit_captures(source, false, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "function_definition", source, start,
                MatchOptions::default(), builder));
        let first = captures(build(0).unwrap().into_typed_pairs());
        let second = captures(build(source.find("int g").unwrap()).unwrap().into_typed_pairs());
        assert!(build(source.find("int broken").unwrap()).is_err());
        for pairs in [first, second] {
            assert!(take_index_region(pairs.iter().find(|pair| pair.as_rule() == Rule::index_suffix).unwrap()).is_some());
            assert!(take_member_region(pairs.iter().find(|pair| pair.as_rule() == Rule::member_suffix).unwrap()).is_some());
        }
    });
    for source in ["int f(){return a[i()].field;}", "int f(){return a[", "int f(){return a."] {
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
                }
                (Err(expected), Err(actual)) => {
                    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                    let values = builder.into_initializers();
                    assert!(values.indices.is_empty() && values.members.is_empty());
                }
                _ => panic!("place hooks changed grammar acceptance: {source}"),
            }
        }
    }
}

#[test]
fn place_regions_preserve_original_index_and_field_anchors_with_replay() {
    let source = "int f(int i){return owner(i)->field[i++].nested + table[(i,2)](i);}";
    with_cache_scope(|| {
        let pairs = build_captures(source, true, |builder|
            super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder))
            .unwrap().into_typed_pairs();
        for pair in captures(pairs) {
            if pair.as_rule() == Rule::index_suffix {
                let child = pair.clone().into_inner().next().unwrap();
                for _ in 0..2 {
                    let region = take_index_region(&pair).unwrap();
                    assert_eq!((region.index.as_rule(), region.index.as_str(), region.index.as_span().start()),
                        (child.as_rule(), child.as_str(), child.as_span().start()));
                    let value = region.finish(Expression::new(ExprKind::Ident("object".into())),
                        Expression::new(ExprKind::Ident("lowered_index".into())));
                    assert!(matches!(value.kind, ExprKind::Index { null_safe: false, .. }));
                }
                assert!(take_member_region(&pair).is_none());
            } else if matches!(pair.as_rule(), Rule::member_suffix | Rule::arrow_suffix) {
                let child = pair.clone().into_inner().next().unwrap();
                for _ in 0..2 {
                    let region = take_member_region(&pair).unwrap();
                    assert_eq!((region.field.as_rule(), region.field.as_str(), region.field.as_span().start()),
                        (child.as_rule(), child.as_str(), child.as_span().start()));
                    let value = region.finish(Expression::new(ExprKind::Ident("lowered_owner".into())));
                    assert!(matches!(value.kind, ExprKind::Member { field, null_safe: false, .. } if field == child.as_str()));
                }
                assert!(take_index_region(&pair).is_none());
            }
        }
    });
}
