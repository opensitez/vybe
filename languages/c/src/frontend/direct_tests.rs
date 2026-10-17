use super::*;
use grameo::{CompileOptions, CompiledGrammar, MatchOptions, ParseError, engine};
use std::sync::LazyLock;

#[cfg(unix)]
fn cpu_ns() -> u128 {
    let mut time = libc::timespec { tv_sec: 0, tv_nsec: 0 };
    assert_eq!(unsafe { libc::clock_gettime(libc::CLOCK_THREAD_CPUTIME_ID, &mut time) }, 0);
    time.tv_sec as u128 * 1_000_000_000 + time.tv_nsec as u128
}

#[cfg(unix)]
#[test]
#[ignore = "paired reference/reduced Direct allocation strategy, same generated matcher and public frontend"]
fn direct_initializer_construction_cost_cpu() {
    let path = std::env::var("VYBE_C_FRONTEND_PREPROCESSED").unwrap();
    let source = std::fs::read_to_string(&path).unwrap();
    let samples: usize = std::env::var("VYBE_C_DIRECT_SAMPLES")
        .ok().map(|s| s.parse().unwrap()).unwrap_or(3);
    assert!(samples > 0);
    eprintln!("direct_cost_layout ast_node_bytes={} compact_frame_bytes={} side_undo_bytes={}",
        std::mem::size_of::<Node>(), std::mem::size_of::<Frame>(),
        std::mem::size_of::<(usize, StateUndo)>());
    let expected = format!("{:?}", crate::parse_with_peg(&source).unwrap());
    for reference in [true, false] {
        let actual = with_reference_construction(reference, || crate::parse_optimized(&source)).unwrap();
        assert_eq!(format!("{actual:?}"), expected, "reference={reference}");
    }
    drop(expected);
    for frontend in [false, true] {
        let run = |reference| with_reference_construction(reference, || {
            if frontend {
                drop(std::hint::black_box(crate::parse_optimized(&source).unwrap()));
                0
            } else {
                let mut builder = Direct::with_capture_elision(&source);
                let tree = super::super::build_hybrid_inner(Rule::program, &source,
                    MatchOptions { max_steps: usize::MAX, ..Default::default() }, &mut builder).unwrap();
                let values = builder.into_initializers();
                let count = values.len();
                drop(std::hint::black_box((tree, values)));
                count
            }
        });
        let count = run(true);
        assert_eq!(run(false), count);
        let mut timings = [Vec::new(), Vec::new()];
        for sample in 0..samples {
            for slot in 0..2 {
                let mode = (sample + slot) % 2;
                let wall = std::time::Instant::now();
                let cpu = cpu_ns();
                assert_eq!(run(mode == 0), count);
                timings[mode].push(cpu_ns() - cpu);
                eprintln!("direct_cost input={path} frontend={frontend} sample={sample} reference={} initializers={count} cpu_ns={} wall_ns={}",
                    mode == 0, timings[mode].last().unwrap(), wall.elapsed().as_nanos());
            }
        }
        for mode in 0..2 {
            timings[mode].sort_unstable();
            eprintln!("direct_cost_median input={path} frontend={frontend} reference={} cpu_ns={} samples={samples}",
                mode == 0, timings[mode][samples / 2]);
        }
    }
}

#[test]
fn direct_initializer_reference_allocations_preserve_owned_values_and_failures() {
    fn receipt(source: &str, reference: bool) -> Result<Vec<((usize, usize), String)>, String> {
        with_reference_construction(reference, || {
            let (_, values) = match_program(source, true, MatchOptions::default())
                .map_err(|error| format!("{error:?}"))?;
            let mut entries: Vec<_> = values.entries.into_iter().map(|(span, cached)|
                (span, format!("{:?}", cached.first.unwrap()))).collect();
            entries.sort_by_key(|entry| entry.0);
            Ok(entries)
        })
    }
    for source in ["int a[]={1,2,3};", "int a[2][2]={{-1,+2},{~3,4}};",
        "int a[2][2]={{1,2},{f(),3}};", "int a[]={-1,~2}; int f( {",
        "int a[]={{1,},,2};"]
    {
        assert_eq!(receipt(source, true), receipt(source, false), "{source}");
    }
}

#[test]
fn direct_initializer_compact_frames_do_not_allocate_ast_slots_for_syntax() {
    let source = "int a[2][2]={{-1,+2},{~3,4}};";
    let run = |reference| with_reference_construction(reference, || {
        let mut builder = Direct::with_capture_elision(source);
        engine::build_capture_program(&*GRAMMAR, "program", source,
            MatchOptions::default(), &mut builder).unwrap();
        assert!(builder.frames.is_empty());
        assert!(builder.state_undo.is_empty());
        let nodes: Vec<_> = builder.arena.into_live_nodes().collect();
        let value = nodes.iter().filter_map(|node| match node {
            Node::Ready { origin: Origin { anchor: Some(span) }, value, .. } => Some((span.end - span.start, value)),
            _ => None,
        }).max_by_key(|(length, _)| *length).map(|(_, value)| value).expect("owned initializer root");
        let ast = format!("{value:?}");
        (nodes.len(), ast)
    });
    let (reference_nodes, reference) = run(true);
    let (compact_nodes, compact) = run(false);
    assert!(compact_nodes < reference_nodes,
        "syntax frame slots were not eliminated: {compact_nodes} vs {reference_nodes}");
    assert_eq!(compact, reference);
}

#[cfg(unix)]
#[test]
#[ignore = "complete Native frontend Direct off/on CPU and PEG AST parity; set VYBE_C_FRONTEND_PREPROCESSED"]
fn direct_initializer_full_frontend_cpu() {
    let path = std::env::var("VYBE_C_FRONTEND_PREPROCESSED").unwrap();
    let source = std::fs::read_to_string(&path).unwrap();
    let samples: usize = std::env::var("VYBE_C_DIRECT_SAMPLES")
        .ok().map(|s| s.parse().unwrap()).unwrap_or(3);
    assert!(samples > 0);
    let run = |enabled: bool| {
        struct Restore(bool);
        impl Drop for Restore {
            fn drop(&mut self) { super::super::DIRECT_DISABLED.with(|s| s.set(self.0)); }
        }
        let _restore = Restore(super::super::DIRECT_DISABLED.with(|s| s.replace(!enabled)));
        crate::parse_optimized(&source).unwrap()
    };
    let expected = format!("{:?}", crate::parse_with_peg(&source).unwrap());
    for enabled in [false, true] {
        let (module, profile) = super::super::profile_construction(|| run(enabled));
        assert_eq!(format!("{module:?}"), expected, "enabled={enabled}");
        if !enabled { assert_eq!(profile.direct_initializer_handoffs, 0); }
        eprintln!("direct_frontend_parity input={path} enabled={enabled} {profile:?}");
    }
    drop(expected);
    let mut timings = [Vec::new(), Vec::new()];
    for sample in 0..samples {
        for slot in 0..2 {
            let mode = (sample + slot) % 2;
            let wall = std::time::Instant::now();
            let cpu = cpu_ns();
            drop(std::hint::black_box(run(mode == 1)));
            timings[mode].push(cpu_ns() - cpu);
            eprintln!("direct_frontend input={path} sample={sample} enabled={} cpu_ns={} wall_ns={}",
                mode == 1, timings[mode].last().unwrap(), wall.elapsed().as_nanos());
        }
    }
    for mode in 0..2 {
        timings[mode].sort_unstable();
        eprintln!("direct_frontend_median enabled={} cpu_ns={} samples={samples}",
            mode == 1, timings[mode][samples / 2]);
    }
}

#[cfg(unix)]
#[test]
#[ignore = "isolated same-generated-matcher Direct off/retained/elided CPU; set VYBE_C_FRONTEND_PREPROCESSED"]
fn direct_initializer_same_match_cpu() {
    let path = std::env::var("VYBE_C_FRONTEND_PREPROCESSED").unwrap();
    let source = std::fs::read_to_string(&path).unwrap();
    let samples: usize = std::env::var("VYBE_C_DIRECT_SAMPLES")
        .ok().map(|s| s.parse().unwrap()).unwrap_or(3);
    assert!(samples > 0);
    let run = |mode: usize| {
        let options = MatchOptions { max_steps: usize::MAX, ..Default::default() };
        if mode == 0 {
            let tree = super::super::build_hybrid_inner(Rule::program, &source, options,
                &mut super::super::NativeBuilder).unwrap();
            drop(tree);
            0
        } else {
            let mut builder = if mode == 2 { Direct::with_capture_elision(&source) }
                else { Direct::new(&source) };
            let tree = super::super::build_hybrid_inner(Rule::program, &source, options,
                &mut builder).unwrap();
            let values = builder.into_initializers();
            let count = values.len();
            drop(tree);
            drop(values);
            count
        }
    };
    let mut counts = [0; 3];
    for mode in 0..3 { counts[mode] = run(mode); }
    assert_eq!(counts[1], counts[2], "elision must not change committed coverage");
    let mut timings = [Vec::new(), Vec::new(), Vec::new()];
    for sample in 0..samples {
        for slot in 0..3 {
            let mode = (sample + slot) % 3;
            let wall = std::time::Instant::now();
            let start = cpu_ns();
            assert_eq!(run(mode), counts[mode]);
            timings[mode].push(cpu_ns() - start);
            eprintln!("direct_same_match input={path} bytes={} sample={sample} mode={} initializers={} cpu_ns={} wall_ns={}",
                source.len(), ["off", "retained", "elided"][mode], counts[mode],
                timings[mode].last().unwrap(), wall.elapsed().as_nanos());
        }
    }
    for mode in 0..3 {
        timings[mode].sort_unstable();
        eprintln!("direct_same_match_median mode={} cpu_ns={} samples={samples}",
            ["off", "retained", "elided"][mode], timings[mode][samples / 2]);
    }
}

static GRAMMAR: LazyLock<CompiledGrammar> = LazyLock::new(|| {
    let source = include_str!("../grammar.pest");
    let source = if source.lines().any(|line|
        line.trim() == "// @grameo owned_capture_boundary initializer_list")
    { source.to_string() } else {
        format!("// @grameo owned_capture_boundary initializer_list\n{source}")
    };
    grameo::compile_with_options(&source, CompileOptions {
        optimizations: true, ..Default::default()
    }).unwrap()
});

fn match_program<'a>(source: &'a str, elide: bool, options: MatchOptions)
    -> Result<(grameo::compat::Pairs<'a, Rule>, Initializers), ParseError>
{
    let mut builder = if elide { Direct::with_capture_elision(source) }
        else { Direct::new(source) };
    let tree = engine::build_capture_program(&*GRAMMAR, "program", source, options, &mut builder)?;
    Ok((tree.into_typed_pairs(), builder.into_initializers()))
}

fn lists<'a>(pairs: grameo::compat::Pairs<'a, Rule>) -> Vec<Pair<'a, Rule>> {
    fn collect<'a>(pair: Pair<'a, Rule>, out: &mut Vec<Pair<'a, Rule>>) {
        if pair.as_rule() == Rule::initializer_list { out.push(pair.clone()); }
        for child in pair.into_inner() { collect(child, out); }
    }
    let mut result = Vec::new();
    for pair in pairs { collect(pair, &mut result); }
    result
}

fn all_captures<'a>(pairs: grameo::compat::Pairs<'a, Rule>) -> Vec<Pair<'a, Rule>> {
    pairs.flat_map(|pair| std::iter::once(pair.clone()).chain(pair.into_inner().flatten())).collect()
}

fn direct_pratt_captures(source: &str) -> grameo::compat::Pairs<'_, Rule> {
    build_captures(source, false, |direct| {
        let mut pratt = super::super::owned::SourcePrattBuilder::native(source);
        let mut builder = super::super::WithDirect::new(&mut pratt, direct);
        super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), &mut builder)
    }).unwrap().into_typed_pairs()
}

#[test]
fn direct_binary_regions_use_actual_pratt_reductions_and_ordered_owned_operands() {
    let source = "int f(int a,int b,int c){return a+b*c;}";
    with_cache_scope(|| {
        let captures = all_captures(direct_pratt_captures(source));
        let pair = captures.iter().find(|pair| pair.as_rule() == Rule::binary_expression && pair.as_str() == "a+b*c").unwrap();
        let name = |name: &str| Expression::new(ExprKind::Ident(name.into()));
        let expected = Expression::new(ExprKind::Binary { op: BinOp::Add,
            left: Box::new(name("a")), right: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Mul, left: Box::new(name("b")), right: Box::new(name("c")),
            })),
        });
        for _ in 0..2 {
            let region = take_binary_region(pair).expect("actual matcher-owned binary AST");
            assert_eq!(region.operand_count(), 3);
            assert_eq!(format!("{:?}", region.value), format!("{expected:?}"));
            let mut order = Vec::new();
            let value = region.lower(|operand, raw| {
                order.push(operand.as_str().to_string());
                with_cache_scope(|| assert!(take_binary_region(pair).is_none()));
                raw
            });
            assert_eq!(order, ["a", "b", "c"]);
            assert_eq!(format!("{value:?}"), format!("{expected:?}"));
        }
        let leaf = captures.iter().find(|pair| pair.as_rule() == Rule::primary_expression && pair.as_str() == "b").unwrap();
        assert!(take_primary_expression(leaf).is_some(), "region ownership lost a contextual leaf handoff");
        let expression = captures.iter().find(|pair| pair.as_rule() == Rule::expression).unwrap();
        assert!(take_syntax_expression(expression).is_none(), "Binary escaped into generic syntax normalization");
    });
}

#[test]
fn direct_binary_regions_preserve_nested_boundaries_and_decline_effectful_operands() {
    for expression in ["a+f()", "a+(b*c)", "p->value+1", "a++ +b", "a+!b", "a+(int)b"] {
        let source = format!("int f(void){{return {expression};}}");
        with_cache_scope(|| {
            let captures = all_captures(direct_pratt_captures(&source));
            let pair = captures.iter().find(|pair| pair.as_rule() == Rule::binary_expression && pair.as_str() == expression).unwrap();
            assert!(take_binary_region(pair).is_none(), "unsupported operand bypassed typed lowering: {expression}");
            if expression == "a+(b*c)" {
                let inner = captures.iter().find(|pair| pair.as_rule() == Rule::binary_expression && pair.as_str() == "b*c").unwrap();
                assert!(take_binary_region(inner).is_some(), "nested grammar region must remain independently consumable");
            }
        });
    }
}

#[test]
fn direct_binary_region_defers_constant_fold_until_after_operand_normalization() {
    let source = "int f(void){return 1==1;}";
    with_cache_scope(|| {
        let captures = all_captures(direct_pratt_captures(source));
        let pair = captures.iter().find(|pair| pair.as_rule() == Rule::binary_expression && pair.as_str() == "1==1").unwrap();
        let region = take_binary_region(pair).unwrap();
        assert!(matches!(region.value.kind, ExprKind::Binary { .. }), "matching folded before contextual operands were normalized");
        let mut index = 0;
        let value = region.lower(|_, raw| {
            index += 1;
            if index == 1 { crate::walker::construct_literal(Rule::int_literal, "2") } else { raw }
        });
        assert_eq!(index, 2);
        assert!(matches!(value.kind, ExprKind::Lit(vybe_ast::Literal::Int(0))));
        let value = take_binary_region(pair).unwrap().lower(|_, raw| raw);
        assert!(matches!(value.kind, ExprKind::Lit(vybe_ast::Literal::Int(1))), "replay retained a context-bound folded result");
    });
}

#[test]
fn direct_binary_reduction_inverse_restores_original_unary_nodes() {
    let source = "a+b";
    let mut builder = Direct::new(source);
    builder.begin(Rule::binary_expression as usize, 0);
    for span in [Span { start: 0, end: 1 }, Span { start: 2, end: 3 }] {
        let operand = Rc::new(SyntaxRecipe::Primary { span, recipe: Recipe::Ident(span) });
        builder.push(Node::Syntax { value: operand.replay(source),
            recipe: Rc::new(SyntaxRecipe::Unary { span, op: None, operand, capture: None }) });
    }
    let mark = builder.checkpoint();
    builder.reduce(Rule::binary_expression as usize, grameo::pratt::Fixity::InfixLeft, Span { start: 1, end: 2 });
    assert!(matches!(builder.arena.get(*builder.arena.values().last().unwrap()), Node::Binary { span: None, .. }));
    builder.rollback(mark);
    assert_eq!(builder.current_frame(), Some((Rule::binary_expression as usize, 0)));
    assert_eq!(builder.arena.values().len(), 2);
    for id in builder.arena.values() { assert!(matches!(builder.arena.get(*id), Node::Syntax { .. })); }
}

#[test]
fn direct_binary_region_validates_all_handles_before_spending_owned_ast() {
    let source = "int f(int a){return a+1;}int g(int b){return b+2;}";
    with_cache_scope(|| {
        let captures = all_captures(direct_pratt_captures(source));
        let pair = captures.iter().find(|pair| pair.as_rule() == Rule::binary_expression && pair.as_str() == "a+1").unwrap();
        let other = captures.iter().find(|pair| pair.as_rule() == Rule::binary_expression && pair.as_str() == "b+2").unwrap();
        let key = (pair.as_span().start(), pair.as_span().end());
        let saved = CACHE.with(|slot| {
            let mut slot = slot.borrow_mut();
            let cache = slot.as_mut().unwrap().get_mut(&(source.as_ptr() as usize, source.len())).unwrap();
            let foreign = cache.binaries.get(&(other.as_span().start(), other.as_span().end())).unwrap().recipe.clone();
            let entry = cache.binaries.get_mut(&key).unwrap();
            std::mem::replace(&mut entry.recipe, foreign)
        });
        assert!(take_binary_region(pair).is_none(), "foreign region handles resolved under the wrong anchor");
        CACHE.with(|slot| {
            let mut slot = slot.borrow_mut();
            let cache = slot.as_mut().unwrap().get_mut(&(source.as_ptr() as usize, source.len())).unwrap();
            let entry = cache.binaries.get_mut(&key).unwrap();
            assert!(entry.first.is_some(), "failed validation spent the first owned AST");
            entry.recipe = saved;
        });
        assert!(take_binary_region(pair).is_some());
        let invalid = "int f(void){return 1+2;} @";
        assert!(build_captures(invalid, false, |direct| {
            let mut pratt = super::super::owned::SourcePrattBuilder::native(invalid);
            let mut builder = super::super::WithDirect::new(&mut pratt, direct);
            super::super::build_hybrid_inner(Rule::program, invalid, MatchOptions::default(), &mut builder)
        }).is_err());
        CACHE.with(|slot| assert!(!slot.borrow().as_ref().unwrap().contains_key(&(invalid.as_ptr() as usize, invalid.len()))));
    });
}

#[test]
fn direct_plain_declarators_own_common_ast_without_type_or_binding_policy() {
    use vybe_ast::BindingPattern;
    let source = "int f(int input){int a,b=17U,c=-input,d=~input;return 0;}";
    with_cache_scope(|| {
        let tree = build_captures(source, false, |builder|
            engine::build_capture_program(&*GRAMMAR, "program", source, MatchOptions::default(), builder)).unwrap();
        let pairs = tree.into_typed_pairs::<Rule>();
        let captures = all_captures(pairs.clone());
        let declarators: Vec<_> = captures.iter().filter(|pair| pair.as_rule() == Rule::init_declarator).collect();
        assert_eq!(declarators.len(), 4);
        for (pair, name) in declarators.into_iter().zip(["a", "b", "c", "d"]) {
            let value = take_plain_declarator(pair).unwrap_or_else(|| panic!("missing owned declaration: {}", pair.as_str()));
            assert!(matches!(&value.declaration.pattern, BindingPattern::Ident(actual) if actual == name));
            assert!(value.declaration.type_hint.is_none());
            assert!(value.declaration.array_bounds.is_none());
            assert!(!value.declaration.with_events);
            assert_eq!(value.syntax.initializer.is_some(), value.declaration.init.is_some(),
                "an incomplete initializer must never authorize a scan bypass");
            let expected = match name {
                "a" => None,
                "b" => Some(crate::walker::construct_literal(Rule::int_literal, "17U")),
                "c" | "d" => Some(crate::walker::construct_numeric_prefix(if name == "c" { "-" } else { "~" },
                    Expression::new(ExprKind::Ident("input".into())))),
                _ => unreachable!(),
            };
            assert_eq!(format!("{:?}", value.declaration.init), format!("{expected:?}"));
            let replay = take_plain_declarator(pair).unwrap();
            assert_eq!(format!("{:?}", replay.declaration), format!("{:?}", value.declaration));
            assert_eq!(replay.syntax, value.syntax);
        }
        let leaf = captures.iter().find(|pair| pair.as_rule() == Rule::primary_expression && pair.as_str() == "input").unwrap();
        assert!(matches!(take_primary_expression(leaf).unwrap().kind, ExprKind::Ident(name) if name == "input"),
            "owned declarations must retain contextual child handoffs");
        let plain = engine::parse_program(&*GRAMMAR, "program", source, MatchOptions::default()).unwrap();
        assert_eq!(capture_shape(pairs.clone().next().unwrap()), capture_shape(plain.into_typed_pairs::<Rule>().next().unwrap()));
    });
}

#[test]
fn direct_plain_declarator_guards_decline_suffixes_and_incomplete_initializers() {
    let source = "void f(void){int *p;int *const q;int a[3];int (*fp)(int);int (group);int b=1+2;int c=fn();int d=(int)1;int e={1};int g=x++;int good=-x;}";
    with_cache_scope(|| {
        let tree = build_captures(source, false, |builder|
            engine::build_capture_program(&*GRAMMAR, "program", source, MatchOptions::default(), builder)).unwrap();
        let pairs = tree.into_typed_pairs::<Rule>();
        let mut accepted = 0;
        for pair in all_captures(pairs.clone()).into_iter().filter(|pair| pair.as_rule() == Rule::init_declarator) {
            if pair.as_str() == "good=-x" {
                let value = take_plain_declarator(&pair).unwrap();
                assert!(value.syntax.initializer.is_some() && value.declaration.init.is_some());
                accepted += 1;
            } else {
                assert!(take_plain_declarator(&pair).is_none(), "unsupported initializer/declarator certified: {}", pair.as_str());
            }
        }
        assert_eq!(accepted, 1);
        let plain = engine::parse_program(&*GRAMMAR, "program", source, MatchOptions::default()).unwrap();
        assert_eq!(capture_shape(pairs.clone().next().unwrap()), capture_shape(plain.into_typed_pairs::<Rule>().next().unwrap()));
    });
}

#[test]
fn direct_plain_declarator_explicit_shapes_keep_generated_capture_and_typed_fallback() {
    for declaration in ["int *p;", "int **p;", "int *const p;", "int *volatile *restrict p;",
        "int (*p);", "int (**p);", "int (p);", "int ((p));", "int p[3];", "int p[];",
        "int p[2][3];", "int *p[3];", "int (*p)[3];", "int p(int);", "int *p(int);",
        "int (*p)(int);", "int (*const p)(int);", "int (*p[2])(int);", "int *p=0;"]
    {
        let mut guard = Direct::new(declaration);
        for rule in [Rule::pointer, Rule::array_suffix, Rule::param_suffix] {
            assert!(guard.observes_rule(rule as usize), "declarator guard must be observed: {rule:?}");
            let mark = guard.checkpoint();
            guard.begin(rule as usize, 0);
            guard.finish(rule as usize, Span { start: 0, end: 0 }, declaration);
            assert!(matches!(guard.arena.get(*guard.arena.values().last().unwrap()), Node::Unsupported));
            guard.rollback(mark);
        }
        with_cache_scope(|| {
            let tree = build_captures(declaration, false, |builder|
                super::super::build_hybrid_inner(Rule::program, declaration, MatchOptions::default(), builder)).unwrap();
            let pairs = tree.into_typed_pairs::<Rule>();
            let declarators: Vec<_> = all_captures(pairs.clone()).into_iter()
                .filter(|pair| pair.as_rule() == Rule::init_declarator).collect();
            assert!(!declarators.is_empty(), "missing explicit declarator control: {declaration}");
            for declarator in declarators {
                assert!(take_plain_declarator(&declarator).is_none(), "explicit shape certified plain: {declaration}");
            }
            let mut native = super::super::NativeBuilder;
            let control = super::super::build_hybrid_inner(Rule::program, declaration,
                MatchOptions::default(), &mut native).unwrap();
            assert_eq!(capture_shape(pairs.clone().next().unwrap()),
                capture_shape(control.into_typed_pairs::<Rule>().next().unwrap()), "{declaration}");
        });
        let expected = crate::parse_with_peg(declaration).unwrap();
        let (actual, counts) = super::super::profile_construction(|| crate::parse_optimized(declaration));
        assert_eq!(format!("{:?}", actual.unwrap()), format!("{expected:?}"), "{declaration}");
        assert_eq!(counts.direct_declaration_handoffs, 0, "explicit shape bypassed typed fallback: {declaration}");
    }
}

#[test]
fn direct_declarator_span_certification_does_not_depend_on_pointer_guard_children() {
    for (source, start, end, plain) in [("name", 0, 4, true), ("name /*tail*/", 0, 4, true),
        ("*name", 1, 5, false), ("(*name)", 2, 6, false), ("(name)", 1, 5, false),
        ("name[2]", 0, 4, false), ("name(int)", 0, 4, false)]
    {
        for reference in [false, true] {
            with_reference_construction(reference, || {
                let mut builder = Direct::new(source);
                builder.begin(Rule::declarator as usize, 0);
                // Deliberately omit pointer/suffix guards: source-span proof
                // must independently reject this hypothetical missing child.
                let span = Span { start, end };
                builder.push(Node::PlainName { span, name: span });
                builder.finish(Rule::declarator as usize, Span { start: 0, end: source.len() }, source);
                assert_eq!(matches!(builder.arena.get(*builder.arena.values().last().unwrap()), Node::PlainName { .. }),
                    plain, "span certificate reference={reference}: {source}");
            });
        }
    }
}

#[test]
fn direct_plain_declarator_units_replay_and_failed_reparse_restore_no_old_values() {
    let source = "int a=x; int b=-y; int broken=;";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start, options| build_unit_captures(source, false, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "declaration", source, start, options, builder));
        let first = all_captures(build(0, MatchOptions::default()).unwrap().into_typed_pairs())
            .into_iter().find(|pair| pair.as_rule() == Rule::init_declarator).unwrap();
        build(source.find("int b").unwrap(), MatchOptions::default()).unwrap();
        assert!(build(source.find("int broken").unwrap(), MatchOptions::default()).is_err());
        let value = take_plain_declarator(&first).unwrap();
        with_cache_scope(|| assert!(take_plain_declarator(&first).is_none()));
        assert_eq!(format!("{:?}", take_plain_declarator(&first).unwrap().declaration), format!("{:?}", value.declaration));
        assert!(build_captures(source, false, |builder|
            engine::build_capture_program(&*GRAMMAR, "program", source, MatchOptions::default(), builder)).is_err());
        assert!(take_plain_declarator(&first).is_none());
    });
}

#[test]
fn direct_plain_declarator_inverse_restores_name_and_owned_initializer() {
    for reference in [false, true] {
        with_reference_construction(reference, || {
            let source = "a=x";
            let mut builder = Direct::new(source);
            builder.begin(Rule::init_declarator as usize, 0);
            builder.push(Node::PlainName { span: Span { start: 0, end: 1 }, name: Span { start: 0, end: 1 } });
            builder.push(Node::Syntax { value: Expression::new(ExprKind::Ident("x".into())),
                recipe: Rc::new(SyntaxRecipe::Initializer { span: Span { start: 2, end: 3 },
                    operand: Rc::new(SyntaxRecipe::Primary { span: Span { start: 2, end: 3 },
                        recipe: Recipe::Ident(Span { start: 2, end: 3 }) }) }) });
            let mark = builder.checkpoint();
            builder.finish(Rule::init_declarator as usize, Span { start: 0, end: 3 }, source);
            assert!(matches!(builder.arena.get(*builder.arena.values().last().unwrap()), Node::Declarator { .. }));
            builder.rollback(mark);
            assert_eq!(builder.current_frame(), Some((Rule::init_declarator as usize, 0)));
            assert!(matches!(builder.arena.get(*builder.arena.values().last().unwrap()), Node::Syntax { .. }));
            builder.finish(Rule::init_declarator as usize, Span { start: 0, end: 3 }, source);
            builder.committed(source, 3);
            assert_eq!(builder.into_initializers().declarators.len(), 1);
        });
    }
}

#[test]
fn direct_plain_declarator_public_backends_execute_identical_standard_wasm() {
    let source = include_str!("../../../../tests/c/doom_startup/frontend_plain_declarator_handoffs.c");
    crate::register();
    let expected = format!("{:?}", crate::parse_with_peg(source).unwrap());
    let mut baseline = None;
    let mut native_handoffs = 0;
    for (index, parse) in [crate::parse_with_peg, crate::parse_optimized, crate::parse_owned_ast,
        crate::parse_owned_contextual, crate::parse_contextual_pratt].into_iter().enumerate()
    {
        let (module, counts) = super::super::profile_construction(|| parse(source));
        let module = module.unwrap();
        if index == 1 { native_handoffs = counts.direct_declaration_handoffs; }
        assert_eq!(format!("{module:?}"), expected, "backend {index}");
        let profile = vybe_compiler::profile::parse_profile(crate::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile).compile(&module).unwrap();
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata: false, include_name_section: true });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes).unwrap();
        if let Some(expected) = &baseline { assert_eq!(&bytes, expected, "backend {index}"); }
        else { baseline = Some(bytes); }
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        vm.run(chunks).unwrap();
        assert_eq!(vm.pending_exit_code, 0, "backend {index}");
    }
    assert!(native_handoffs > 0, "the public Native path must consume actual grammar-built declarations");
}

#[test]
fn direct_primary_syntax_preserves_unresolved_names_and_literal_policy() {
    let source = "int f(int x){int y=17U; x; __func__; NULL; stdin; environ; ATOMIC_FLAG_INIT; (x); x+2; return y;}";
    with_cache_scope(|| {
        let tree = build_captures(source, true, |builder|
            engine::build_capture_program(&*GRAMMAR, "program", source, MatchOptions::default(), builder)).unwrap();
        let pairs = tree.into_typed_pairs::<Rule>();
        let primaries: Vec<_> = all_captures(pairs.clone()).into_iter()
            .filter(|pair| pair.as_rule() == Rule::primary_expression).collect();
        for pair in &primaries {
            if pair.as_str() == "(x)" {
                assert!(take_primary_expression(pair).is_none(), "parenthesis must retain its expression barrier");
                continue;
            }
            let value = take_primary_expression(pair).unwrap_or_else(|| panic!("missing primary {}", pair.as_str()));
            let expected = match pair.as_str() {
                "17U" => crate::walker::construct_literal(Rule::int_literal, "17U"),
                "2" => crate::walker::construct_literal(Rule::int_literal, "2"),
                name => Expression::new(ExprKind::Ident(name.to_string())),
            };
            assert_eq!(format!("{value:?}"), format!("{expected:?}"));
            assert_eq!(format!("{:?}", take_primary_expression(pair).unwrap()), format!("{expected:?}"),
                "predeclaration replay must preserve unresolved syntax");
            assert!(take_primary_expression(&pair.clone().into_inner().next().unwrap()).is_none(),
                "handoff requires the exact primary rule, not just an overlapping span");
        }
        let plain = engine::parse_program(&*GRAMMAR, "program", source, MatchOptions::default()).unwrap();
        assert_eq!(capture_shape(pairs.clone().next().unwrap()),
            capture_shape(plain.into_typed_pairs::<Rule>().next().unwrap()));
    });
}

#[test]
fn direct_expression_hooks_construct_numeric_prefixes_and_ordered_comma_syntax() {
    let source = "void f(void){-x,+y,~3;}";
    with_cache_scope(|| {
        let tree = build_captures(source, true, |builder|
            engine::build_capture_program(&*GRAMMAR, "program", source, MatchOptions::default(), builder)).unwrap();
        let captures = all_captures(tree.into_typed_pairs::<Rule>());
        let expression = captures.iter().find(|pair| pair.as_rule() == Rule::expression).unwrap();
        let name = |name: &str| Expression::new(ExprKind::Ident(name.to_string()));
        let expected = Expression::new(ExprKind::Sequence(vec![
            crate::walker::construct_numeric_prefix("-", name("x")), name("y"),
            crate::walker::construct_numeric_prefix("~", crate::walker::construct_literal(Rule::int_literal, "3")),
        ]));
        assert_eq!(format!("{:?}", take_syntax_expression(expression).unwrap()), format!("{expected:?}"));
        assert_eq!(format!("{:?}", take_syntax_expression(expression).unwrap()), format!("{expected:?}"));
        let operand = captures.iter().find(|pair| pair.as_rule() == Rule::primary_expression && pair.as_str() == "x").unwrap();
        assert_eq!(format!("{:?}", take_primary_expression(operand).unwrap()), format!("{:?}", name("x")),
            "moving an enclosing syntax tree must preserve predeclaration's leaf handoffs");
        let unary = captures.iter().find(|pair| pair.as_rule() == Rule::unary_expression && pair.as_str() == "-x").unwrap();
        assert_eq!(format!("{:?}", take_syntax_expression(unary).unwrap()),
            format!("{:?}", crate::walker::construct_numeric_prefix("-", name("x"))));
    });
}

#[test]
fn direct_grouped_syntax_moves_compound_ast_without_widening_leaf_primary_api() {
    let source = "void f(void){-((x,+y)),~(3);}";
    with_cache_scope(|| {
        let tree = build_captures(source, false, |builder|
            super::super::build_hybrid_inner(Rule::program, source, MatchOptions::default(), builder)).unwrap();
        let pairs = tree.into_typed_pairs::<Rule>();
        let captures = all_captures(pairs.clone());
        let root = captures.iter().find(|pair| pair.as_rule() == Rule::expression).unwrap();
        let name = |name: &str| Expression::new(ExprKind::Ident(name.into()));
        let expected = Expression::new(ExprKind::Sequence(vec![
            crate::walker::construct_numeric_prefix("-", Expression::new(ExprKind::Sequence(vec![name("x"), name("y")]))),
            crate::walker::construct_numeric_prefix("~", crate::walker::construct_literal(Rule::int_literal, "3")),
        ]));
        for _ in 0..2 {
            assert_eq!(format!("{:?}", take_syntax_expression(root).unwrap()), format!("{expected:?}"));
        }
        for primary in captures.iter().filter(|pair| pair.as_rule() == Rule::primary_expression) {
            if primary.as_str().starts_with('(') {
                assert!(take_primary_expression(primary).is_none(), "leaf-only API admitted compound syntax");
            } else {
                assert!(take_primary_expression(primary).is_some(), "enclosing ownership lost its leaf handoff");
            }
        }
        let control = engine::parse_program(&*GRAMMAR, "program", source, MatchOptions::default()).unwrap();
        assert_eq!(capture_shape(pairs.clone().next().unwrap()), capture_shape(control.into_typed_pairs::<Rule>().next().unwrap()));
    });
}

#[test]
fn direct_grouped_syntax_preserves_full_typed_ast_and_contextual_fallback() {
    for source in [
        "int f(int x,int y){int a=-(x),b=~((y)),c=(x,y);return -(a);}",
        "int f(void){return (__func__,0);}",
        "typedef int *P; int f(P p){return (p)==0;}",
        "int f(void){int x=1;{int x=2,y=(x);if(y!=2)return 1;}return x;}",
        "int f(int a,int b){return (a+b);}",
        "int f(int *p){return *(p);}",
        "int f(int x){return (x++);}",
        "int f(int x){return ((int)x);}",
    ] {
        let expected = crate::parse_with_peg(source).unwrap();
        let (actual, counts) = super::super::profile_construction(|| crate::parse_optimized(source));
        assert_eq!(format!("{:?}", actual.unwrap()), format!("{expected:?}"), "{source}");
        assert!(counts.direct_primary_handoffs + counts.direct_syntax_handoffs
            + counts.direct_binary_handoffs > 0, "{source}");
    }
}

#[test]
fn direct_grouped_syntax_budget_and_late_failure_publish_no_partial_trees() {
    for source in ["int f(int x){int y=-((x));return ~(y);}",
        "int f(int x){int y=-((x));return ~(y);} @", "int f(int x){return -(x;}"]
    {
        for max_steps in [1, 200, usize::MAX] {
            let options = MatchOptions { max_steps, ..Default::default() };
            let control = engine::parse_program(&*GRAMMAR, "program", source, options.clone());
            let mut builder = Direct::new(source);
            let candidate = engine::build_capture_program(&*GRAMMAR, "program", source, options, &mut builder);
            match (control, candidate) {
                (Ok(_), Ok(_)) => assert!(!builder.into_initializers().syntax.is_empty()),
                (Err(expected), Err(actual)) => {
                    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                    let cache = builder.into_initializers();
                    assert!(cache.syntax.is_empty() && cache.primaries.is_empty() && cache.declarators.is_empty());
                }
                _ => panic!("grouped syntax acceptance changed: {source}"),
            }
        }
    }
}

#[test]
fn direct_expression_guards_keep_contextual_forms_and_exact_capture_shape() {
    for expression in ["x+1", "f(x)", "x=1", "!x", "*x", "&x", "++x", "x++",
        "(int)x", "sizeof(x)", "_Alignof(int)", "x?y:z", "x?:y", "({x;})", "(x+1)",
        "-f(x)", "~(int)x", "x, f(x)", "x[1]", "x.member", "x->member"]
    {
        let source = format!("void f(void){{{expression};}}");
        with_cache_scope(|| {
            let tree = build_captures(&source, true, |builder|
                engine::build_capture_program(&*GRAMMAR, "program", &source, MatchOptions::default(), builder)).unwrap();
            let pairs = tree.into_typed_pairs::<Rule>();
            let captures = all_captures(pairs.clone());
            let root = captures.iter().find(|pair| pair.as_rule() == Rule::expression && pair.as_str() == expression).unwrap();
            assert!(take_syntax_expression(root).is_none(), "unsupported syntax lost its guard: {expression}");
            let control = engine::parse_program(&*GRAMMAR, "program", &source, MatchOptions::default()).unwrap();
            assert_eq!(capture_shape(pairs.clone().next().unwrap()),
                capture_shape(control.into_typed_pairs::<Rule>().next().unwrap()), "{expression}");
        });
    }
}

#[test]
fn direct_expression_prefix_and_sequence_inverse_restore_original_owned_nodes() {
    let source = "-name";
    let span = Span { start: 0, end: source.len() };
    let mut builder = Direct::new(source);
    builder.begin(Rule::unary_expression as usize, 0);
    builder.push(Node::Prefix("-"));
    builder.push(Node::Primary { span: Span { start: 1, end: source.len() },
        value: Expression::new(ExprKind::Ident("name".into())),
        recipe: Recipe::Ident(Span { start: 1, end: source.len() }) });
    let mark = builder.checkpoint();
    builder.finish(Rule::unary_expression as usize, span, source);
    assert!(matches!(builder.arena.get(*builder.arena.values().last().unwrap()), Node::Syntax { .. }));
    builder.rollback(mark);
    assert_eq!(builder.current_frame(), Some((Rule::unary_expression as usize, 0)));
    assert!(matches!(builder.arena.get(*builder.arena.values().last().unwrap()), Node::Primary { .. }));
    builder.finish(Rule::unary_expression as usize, span, source);
    builder.committed(source, source.len());
    assert!(!builder.into_initializers().syntax.is_empty());
}

#[test]
fn direct_primary_unit_commit_and_failed_reparse_obey_source_generation() {
    let source = "int a=x; int b=y; int broken=;";
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start, options| build_unit_captures(source, true, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "declaration", source, start, options, builder));
        let first = all_captures(build(0, MatchOptions::default()).unwrap().into_typed_pairs())
            .into_iter().find(|pair| pair.as_rule() == Rule::primary_expression).unwrap();
        build(source.find("int b").unwrap(), MatchOptions::default()).unwrap();
        assert!(build(source.find("int broken").unwrap(), MatchOptions::default()).is_err());
        assert!(take_primary_expression(&first).is_some());
        with_cache_scope(|| assert!(take_primary_expression(&first).is_none()));
        assert!(take_primary_expression(&first).is_some());
        assert!(build_captures(source, true, |builder|
            engine::build_capture_program(&*GRAMMAR, "program", source, MatchOptions::default(), builder)).is_err());
        assert!(take_primary_expression(&first).is_none(), "failed whole reparse resurrected old primary syntax");
    });
}

#[test]
fn direct_primary_finish_rollback_restores_owned_operand_and_semantic_frames() {
    for reference in [false, true] {
        with_reference_construction(reference, || {
            let source = "name";
            let span = Span { start: 0, end: source.len() };
            let mut builder = Direct::new(source);
            builder.begin(Rule::primary_expression as usize, 0);
            builder.begin(Rule::ident_name as usize, 0);
            builder.finish(Rule::ident_name as usize, span, source);
            let mark = builder.checkpoint();
            builder.finish(Rule::primary_expression as usize, span, source);
            assert!(matches!(builder.arena.get(*builder.arena.values().last().unwrap()), Node::Primary { .. }));
            builder.rollback(mark);
            assert_eq!(builder.current_frame(), Some((Rule::primary_expression as usize, 0)));
            assert!(matches!(builder.arena.get(*builder.arena.values().last().unwrap()),
                Node::Ready { value: Expression { kind: ExprKind::Ident(name), .. }, .. } if name == "name"));
            builder.finish(Rule::primary_expression as usize, span, source);
            builder.committed(source, source.len());
            let mut values = builder.into_initializers();
            let entry = values.primaries.get_mut(&(0, source.len())).unwrap();
            assert_eq!(format!("{:?}", entry.first.take().unwrap()), format!("{:?}", entry.recipe.replay(source)));
        });
    }
}

#[test]
fn direct_body_hooks_construct_controls_and_preserve_ordered_contextual_anchors() {
    let source = "void f(int n){;while(n){if(n)break;continue;}goto done;done:return;}";
    let body_anchor = with_cache_scope(|| {
        let tree = build_captures(source, true, |builder|
            engine::build_capture_program(&*GRAMMAR, "program", source, MatchOptions::default(), builder)).unwrap();
        let pairs = tree.into_typed_pairs::<Rule>();
        let captures = all_captures(pairs.clone());
        let body = captures.iter().find(|pair| pair.as_rule() == Rule::compound_statement).unwrap();
        let items = take_body(body).expect("grammar-finished function body");
        let labels: Vec<_> = items.iter().map(|item| match item {
            BodyItem::Owned { pair, .. } | BodyItem::Contextual(pair) => pair.as_str(),
        }).collect();
        assert_eq!(labels, [";", "while(n){if(n)break;continue;}", "goto done;", "done:return;"]);
        assert!(matches!(&items[0], BodyItem::Owned { value: OwnedStatement::Empty, .. }));
        assert!(matches!(&items[1], BodyItem::Contextual(_)));
        assert!(matches!(&items[2], BodyItem::Owned { value: OwnedStatement::Emit(Statement {
            kind: StmtKind::GoTo(label), .. }), .. } if label == "done"));
        assert!(matches!(&items[3], BodyItem::Contextual(_)));
        let again = take_body(body).unwrap();
        assert!(matches!(&again[2], BodyItem::Owned { value: OwnedStatement::Emit(Statement {
            kind: StmtKind::GoTo(label), .. }), .. } if label == "done"));
        for (text, expected) in [("break;", "break"), ("continue;", "continue"), ("return;", "return")] {
            let pair = captures.iter().find(|pair| pair.as_rule() == Rule::statement && pair.as_str() == text).unwrap();
            let Some(OwnedStatement::Emit(value)) = take_control_statement(pair) else { panic!("{text}"); };
            assert!(match value.kind {
                StmtKind::Break(BreakTarget::Implicit) => expected == "break",
                StmtKind::Continue(ContinueTarget::Implicit) => expected == "continue",
                StmtKind::Return(None) => expected == "return",
                _ => false,
            });
        }
        // Owning a block/control never removes captures needed by metadata or
        // the contextual Walker; only complete initializer islands may elide.
        let plain = engine::parse_program(&*GRAMMAR, "program", source, MatchOptions::default()).unwrap();
        assert_eq!(capture_shape(pairs.clone().next().unwrap()),
            capture_shape(plain.into_typed_pairs::<Rule>().next().unwrap()));
        body.clone()
    });
    assert!(take_body(&body_anchor).is_none());
}

#[test]
fn direct_body_contextual_returns_do_not_bypass_typed_lowering() {
    for source in ["int f(int x){return x;}", "int f(void){return 1;}",
        "int f(void){return (1,2);}", "int f(void){return ({return 1;2;});}",
        "void f(void){continue label;}"]
    {
        let (pairs, mut cache) = match_program(source, true, MatchOptions::default()).unwrap();
        for pair in all_captures(pairs).into_iter().filter(|pair| pair.as_rule() == Rule::statement) {
            if pair.as_str() == "continue label;" {
                assert!(matches!(cache.take_control(&pair), Some(OwnedStatement::Emit(Statement {
                    kind: StmtKind::Continue(ContinueTarget::Label(label)), .. })) if label == "label"));
            } else if pair.as_str().starts_with("return") {
                assert!(cache.take_control(&pair).is_none(), "typed return was preconstructed: {source}");
            }
        }
    }
}

#[test]
fn direct_body_failed_units_and_stale_handles_do_not_spend_owned_values() {
    let source = "void a(void){return;} void b(void){goto done;done:return;} void broken( {";
    let second = source.find("void b").unwrap();
    let failed = source.find("void broken").unwrap();
    with_cache_scope(|| {
        reset_unit_captures(source);
        let build = |start, options| build_unit_captures(source, true, |builder|
            engine::build_capture_program_at(&*GRAMMAR, "function_definition", source, start, options, builder));
        let first = all_captures(build(0, MatchOptions::default()).unwrap().into_typed_pairs());
        let second = all_captures(build(second, MatchOptions::default()).unwrap().into_typed_pairs());
        assert!(build(failed, MatchOptions::default()).is_err());
        assert!(build(0, MatchOptions { max_steps: 1, ..Default::default() }).is_err());
        for captures in [&first, &second] {
            let body = captures.iter().find(|pair| pair.as_rule() == Rule::compound_statement).unwrap();
            assert!(take_body(body).is_some(), "later unit failure invalidated a committed body");
        }
        let first_body = first.iter().find(|pair| pair.as_rule() == Rule::compound_statement).unwrap();
        let second_body = second.iter().find(|pair| pair.as_rule() == Rule::compound_statement).unwrap();
        CACHE.with(|slot| {
            let mut cache = slot.borrow_mut();
            let cache = cache.as_mut().unwrap().get_mut(&(source.as_ptr() as usize, source.len())).unwrap();
            let saved = cache.bodies.get(&(first_body.as_span().start(), first_body.as_span().end())).unwrap().clone();
            let key = (second_body.as_span().start(), second_body.as_span().end());
            let original = cache.bodies.insert(key, saved).unwrap();
            assert!(cache.take_body(second_body).is_none(), "foreign unit handles resolved under another arena");
            cache.bodies.insert(key, original);
        });
        assert!(take_body(second_body).is_some(), "failed handle validation spent owned statements");
        let invalid = "void f(void){goto done;return;} @";
        let mut builder = Direct::with_capture_elision(invalid);
        assert!(engine::build_capture_program(&*GRAMMAR, "program", invalid,
            MatchOptions::default(), &mut builder).is_err());
        let failed = builder.into_initializers();
        assert!(failed.bodies.is_empty() && failed.statements.is_empty());
    });
}

#[derive(Debug, PartialEq, Eq)]
struct CaptureShape<'a> {
    source: &'a str,
    nodes: Vec<(Rule, usize, usize, usize)>,
}

fn capture_shape<'a>(pair: Pair<'a, Rule>) -> CaptureShape<'a> {
    let source = pair.as_span().get_input();
    let mut nodes = Vec::new();
    let mut pending = vec![(pair, 0)];
    while let Some((pair, depth)) = pending.pop() {
        let span = pair.as_span();
        nodes.push((pair.as_rule(), span.start(), span.end(), depth));
        // Preorder plus depth uniquely preserves ordered tree structure, even
        // when several grammar wrappers have identical rules or byte spans.
        let first_child = pending.len();
        pending.extend(pair.into_inner().map(|child| (child, depth + 1)));
        pending[first_child..].reverse();
    }
    CaptureShape { source, nodes }
}

#[test]
fn capture_shape_records_nested_wrappers_once_without_recursive_debug_strings() {
    fn reference(pair: Pair<'_, Rule>, depth: usize, nodes: &mut Vec<(Rule, usize, usize, usize)>) {
        let span = pair.as_span();
        nodes.push((pair.as_rule(), span.start(), span.end(), depth));
        for child in pair.into_inner() { reference(child, depth + 1, nodes); }
    }
    for nesting in [0, 2, 4] {
        let source = format!("int f(int x){{return {}x{};}}", "(".repeat(nesting), ")".repeat(nesting));
        let pairs = engine::parse_program(&*GRAMMAR, "program", &source, MatchOptions::default())
            .unwrap().into_typed_pairs::<Rule>();
        let pair = pairs.clone().next().unwrap();
        let actual = capture_shape(pair.clone());
        let mut expected = Vec::new();
        reference(pair, 0, &mut expected);
        assert_eq!(actual.source, source);
        assert_eq!(actual.nodes, expected);
        assert_eq!(actual.nodes.len(), all_captures(pairs).len());
        assert!(format!("{actual:?}").len() < source.len() * 2 + actual.nodes.len() * 128,
            "flat capture diagnostics must remain linear in source and capture count");
    }
}

fn reference(pair: Pair<'_, Rule>) -> Expression {
    if Direct::literal(pair.as_rule() as usize).is_some() {
        return crate::walker::construct_literal(pair.as_rule(), pair.as_str());
    }
    if pair.as_rule() == Rule::initializer_list {
        return array(pair.into_inner().map(reference));
    }
    if pair.as_rule() == Rule::unary_expression {
        let mut children = pair.clone().into_inner();
        let first = children.next().unwrap();
        if first.as_rule() == Rule::prefix_op {
            let operand = reference(children.next().unwrap());
            assert!(children.next().is_none());
            return crate::walker::construct_numeric_prefix(first.as_str(), operand);
        }
    }
    let mut children = pair.into_inner();
    let value = reference(children.next().expect("reference positional child"));
    assert!(children.next().is_none(), "contextual syntax is not a direct node");
    value
}

#[test]
fn direct_initializer_hooks_move_nested_arrays_and_replay_exact_literal_policy() {
    for source in [
        "int a[]={1,2,3};",
        "int a[]={-1,+2,~3,- -4,+ +1,+(~5),~(-6)};",
        "int a[2][3]={{1,2,3},{4,5,6}};",
        "unsigned long long a[]={0xffffffffu,0x8000000000000000ULL,18446744073709551615ULL};",
        "double a[]={0x1.fp2,1e-3f,((2.5)),3L};",
        "int a[]={'a','\\n','\\x7f','\\123',u8'x',L'x',u'x',U'x',true,false};",
        "char *a[]={\"ab\" \"cd\",u8\"caf\u{e9}\",L\"wide\",u\"wide\",U\"wide\"};",
        "int a[]={/* before */(1), // middle\n {2,3},};",
        "int a[]={};",
    ] {
        let (pairs, mut values) = match_program(source, false, MatchOptions::default()).unwrap();
        let list = lists(pairs).remove(0);
        let expected = reference(list.clone());
        let first = values.take(&list).expect("completed direct initializer");
        let second = values.take(&list).expect("repeat consumption replays syntax");
        assert_eq!(format!("{first:?}"), format!("{expected:?}"), "{source}");
        assert_eq!(format!("{second:?}"), format!("{expected:?}"), "{source}");
    }
}

#[test]
fn direct_initializer_elision_retains_the_anchor_without_constructing_descendant_captures() {
    let source = "int a[2][3]={{1,2,3},{4,5,6}};";
    let (reference_pairs, _) = match_program(source, false, MatchOptions::default()).unwrap();
    let expected = reference(lists(reference_pairs).remove(0));
    let (pairs, mut values) = match_program(source, true, MatchOptions::default()).unwrap();
    let list = lists(pairs).remove(0);
    assert_eq!(list.as_str(), "{{1,2,3},{4,5,6}}");
    assert_eq!(list.clone().into_inner().count(), 0,
        "a certified complete island must omit descendants, not just cache another AST");
    assert_eq!(format!("{:?}", values.take(&list).unwrap()), format!("{expected:?}"));
}

#[test]
fn direct_initializer_contextual_forms_keep_captures_and_order() {
    for source in [
        "int a[]={x,2};", "int a[]={f(1),2};", "int a[]={1+2,3};",
        "int a[]={(1,2),3};", "int a[]={(1)(2),3};", "int a[]={!1,2};",
        "int a[]={-x,2};", "int a[]={++1,2};", "int a[]={--true,2};", "int a[]={-\"x\",2};",
        "int a[]={(int)1,2};", "int a[]={sizeof(int),2};",
        // PEG compatibility controls, not ISO C preprocessing-number claims.
        "int a[]={0x1e-2};", "int a[]={0xbe+2};", "int a[]={0XFE-2};",
        "int a[]={[2]=3,4};", "struct S {int x;}; struct S a={.x=1};",
    ] {
        let (pairs, mut values) = match_program(source, true, MatchOptions::default()).unwrap();
        let list = lists(pairs).remove(0);
        assert!(values.take(&list).is_none(), "contextual initializer admitted: {source}");
        assert!(list.into_inner().count() > 0, "fallback captures lost: {source}");
    }
}

#[test]
fn direct_initializer_numeric_admission_distinguishes_hex_digits_from_exponents() {
    for source in ["{0x1e-2}", "{0xbe+2}", "{0XFE-2}", "{0x1E+2}"] {
        assert_eq!(initializer_end(source, 0), None, "{source}");
    }
    for source in ["{0x1p-2}", "{0X1.EP+2}", "{1e-2}", "{1E+2}", "{.1e-2}"] {
        assert_eq!(initializer_end(source, 0), Some(source.len()), "{source}");
        let source = format!("double a[]={source};");
        let (pairs, mut values) = match_program(&source, true, MatchOptions::default()).unwrap();
        assert!(values.take(&lists(pairs).remove(0)).is_some(), "{source}");
    }
}

#[test]
fn direct_initializer_numeric_prefixes_preserve_full_public_ast_and_width_policy() {
    for source in [
        "int a[]={-1,+2,~3,- -4,+(~5),~(-6)};",
        "unsigned int a[]={-1u,~0xffffffffU,+0x80000000U};",
        "unsigned long long a[]={-1ULL,~0ULL,+0x8000000000000000ULL};",
        "long long a[]={-2147483648LL,~0LL,- -9223372036854775807LL};",
        "double a[]={-0.0,+1e-2,-0X1.EP+2,-(-0x1p-2)};",
        "int a[]={-'a',+u8'x',~L'x',-true,+false};",
    ] {
        let expected = crate::parse_with_peg(source).unwrap();
        for parse in [crate::parse_optimized, crate::parse_contextual_pratt] {
            let (actual, profile) = super::super::profile_construction(|| parse(source));
            assert_eq!(format!("{:?}", actual.unwrap()), format!("{expected:?}"), "{source}");
            assert!(profile.direct_initializer_handoffs > 0, "{source}: {profile:?}");
        }
    }
}

#[test]
fn direct_initializer_numeric_prefix_fixture_executes_identical_standard_wasm() {
    let source = include_str!("../../../../tests/c/doom_startup/frontend_numeric_prefix_handoffs.c");
    crate::register();
    let expected = format!("{:?}", crate::parse_with_peg(source).unwrap());
    let mut baseline = None;
    for parse in [crate::parse_with_peg, crate::parse_optimized, crate::parse_owned_ast,
        crate::parse_owned_contextual, crate::parse_contextual_pratt]
    {
        let module = parse(source).unwrap();
        assert_eq!(format!("{module:?}"), expected);
        let profile = vybe_compiler::profile::parse_profile(crate::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module).unwrap();
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false, include_name_section: true,
            });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap();
        if let Some(expected) = &baseline { assert_eq!(&bytes, expected); }
        else { baseline = Some(bytes); }
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        vm.run(chunks).unwrap();
        assert_eq!(vm.pending_exit_code, 0);
    }
}

#[test]
fn direct_initializer_admission_never_elides_a_contextual_success() {
    let atoms = ["1", "0x1e", "0xbe", "0XFE", "0x1p2", "1e2", "1e-2", "true", "'e'", "\"e\""];
    let separators = ["+", "-", "*", "/", "<<", "==", "?1:", ","];
    for left in atoms {
        for right in atoms {
            for separator in separators {
                let source = format!("int a[]={{({left}{separator}{right})}};");
                let mut reference = super::super::NativeBuilder;
                let expected = engine::build_capture_program(&*GRAMMAR, "program", &source,
                    MatchOptions::default(), &mut reference);
                let actual = match_program(&source, true, MatchOptions::default());
                match (expected, actual) {
                    (Ok(expected), Ok((pairs, mut values))) => {
                        let baseline = lists(expected.into_typed_pairs()).remove(0);
                        let actual = lists(pairs).remove(0);
                        assert!(values.take(&actual).is_none(), "contextual AST admitted: {source}");
                        assert_eq!(capture_shape(actual), capture_shape(baseline), "captures changed: {source}");
                    }
                    (Err(expected), Err(actual)) => assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}"),
                    _ => panic!("acceptance changed: {source}"),
                }
            }
        }
    }
}

#[test]
fn direct_initializer_prefix_admission_preserves_tokens_captures_and_errors() {
    for prefix in ["++", "--", "+ +", "- -", "+ -", "- +", "~", "+ + +",
        "- - -", "!", "~++", "~--", "+/*c*/+", "-/*c*/-"]
    {
        for atom in ["1", "true", "0x1e-2", "\"x\"", "f()", "x"] {
            let source = format!("int a[]={{({prefix}{atom})}};");
            let expected = engine::parse_program(&*GRAMMAR, "program", &source,
                MatchOptions::default());
            let actual = match_program(&source, true, MatchOptions::default());
            match (expected, actual) {
                (Ok(expected), Ok((pairs, mut values))) => {
                    let baseline = lists(expected.into_typed_pairs()).remove(0);
                    let actual = lists(pairs).remove(0);
                    if let Some(value) = values.take(&actual) {
                        assert_eq!(format!("{value:?}"), format!("{:?}", reference(baseline)), "{source}");
                        assert_eq!(actual.into_inner().count(), 0, "{source}");
                    } else {
                        assert_eq!(capture_shape(actual), capture_shape(baseline), "{source}");
                    }
                }
                (Err(expected), Err(actual)) => assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}"),
                _ => panic!("prefix acceptance changed: {source}"),
            }
        }
    }
    for source in ["{++1}", "{--true}", "{~++1}", "{~--1}"] {
        assert_eq!(initializer_end(source, 0), None, "{source}");
    }
}

#[test]
fn direct_initializer_declined_table_does_not_construct_speculative_scalar_asts() {
    let source = "int a[]={1,2,3,!4,5,6};";
    let mut builder = Direct::with_capture_elision(source);
    let tree = engine::build_capture_program(&*GRAMMAR, "program", source,
        MatchOptions::default(), &mut builder).unwrap();
    assert!(lists(tree.into_typed_pairs()).remove(0).into_inner().count() > 0);
    assert_eq!(builder.depth, 0);
    assert!(!builder.arena.into_live_nodes().any(|node| matches!(node, Node::Ready { .. })),
        "a declined initializer must not construct unused positive literal ASTs");
}

#[test]
fn direct_initializer_cache_is_parse_scoped_repeatable_and_failure_invalidated() {
    let source = "int a[]={1,2,3};";
    let build = |source, options| {
        build_captures(source, true, |builder|
            engine::build_capture_program(&*GRAMMAR, "program", source, options, builder))
    };
    with_cache_scope(|| {
        let tree = build(source, MatchOptions::default()).unwrap();
        let list = lists(tree.into_typed_pairs()).remove(0);
        assert!(take_initializer_list(&list).is_some());
        assert!(take_initializer_list(&list).is_some());
        with_cache_scope(|| assert!(take_initializer_list(&list).is_none()));
        assert!(take_initializer_list(&list).is_some(), "outer cache not restored");
        let _ = std::panic::catch_unwind(|| with_cache_scope(|| panic!("nested scope")));
        assert!(take_initializer_list(&list).is_some(), "unwind lost outer cache");
        let mut bounded = MatchOptions::default();
        bounded.max_steps = 1;
        assert!(build(source, bounded).is_err());
        assert!(take_initializer_list(&list).is_none(), "failed reparse retained stale AST");
        let other = source.to_string();
        let (pairs, _) = match_program(&other, false, MatchOptions::default()).unwrap();
        assert!(take_initializer_list(&lists(pairs).remove(0)).is_none(),
            "equal text at a different source identity must not alias");
    });
}

#[test]
fn direct_initializer_absolute_units_append_only_committed_values_and_keep_leases() {
    let source = "int a[2]={1,-2};\nint b[2]={~3,+4};\nint bad[]={5} broken";
    let second = source.find("int b").unwrap();
    let failed = source.find("int bad").unwrap();
    let build = |start, options| build_unit_captures(source, true, |builder|
        engine::build_capture_program_at(&*GRAMMAR, "declaration", source, start, options, builder));
    fn address(value: &Expression) -> usize {
        let ExprKind::Array(values) = &value.kind else { panic!("array unit"); };
        values.as_ptr() as usize
    }
    with_cache_scope(|| {
        reset_unit_captures(source);
        let first = lists(build(0, MatchOptions::default()).unwrap().into_typed_pairs()).remove(0);
        let lease = lease_initializer_list(&first).unwrap();
        let original = address(lease.as_ref());
        let second_list = lists(build(second, MatchOptions::default()).unwrap().into_typed_pairs()).remove(0);
        assert!(second_list.as_span().start() >= second, "unit spans must stay absolute");
        assert_eq!(second_list.clone().into_inner().count(), 0);
        build(0, MatchOptions::default()).unwrap();
        drop(lease);
        assert_eq!(address(&take_initializer_list(&first).unwrap()), original,
            "append/replayed unit replaced a previous unit's leased value");
        let lease = lease_initializer_list(&first).unwrap();
        let before_failure = address(lease.as_ref());
        assert!(build(failed, MatchOptions::default()).is_err());
        assert!(build(second, MatchOptions { max_steps: 1, ..Default::default() }).is_err());
        let uncommitted: Result<(), ()> = build_unit_captures(source, true, |_| Ok(()));
        assert!(uncommitted.is_ok());
        drop(lease);
        assert_eq!(address(&take_initializer_list(&first).unwrap()), before_failure,
            "a failed later unit invalidated an earlier unit's live lease");
        assert!(take_initializer_list(&second_list).is_some());
        CACHE.with(|cache| assert_eq!(cache.borrow().as_ref().unwrap()
            .get(&(source.as_ptr() as usize, source.len())).unwrap().len(), 2,
            "failed/uncommitted units published entries"));
        with_cache_scope(|| {
            reset_unit_captures(source);
            assert!(take_initializer_list(&first).is_none());
        });
        assert!(take_initializer_list(&first).is_some(), "nested sequence lost outer units");
        let old_lease = lease_initializer_list(&first).unwrap();
        reset_unit_captures(source);
        build(second, MatchOptions::default()).unwrap();
        drop(old_lease);
        assert!(take_initializer_list(&first).is_none());
        assert!(take_initializer_list(&second_list).is_some(), "old lease overwrote new sequence");
    });
    let first = lists(build(0, MatchOptions::default()).unwrap().into_typed_pairs()).remove(0);
    assert!(first.into_inner().count() > 0, "outside a cache scope units must retain captures");
}

#[test]
fn direct_initializer_read_lease_keeps_first_ast_and_respects_reentrant_invalidation() {
    let source = "int a[]={1,2,3};";
    let build = |options| build_captures(source, true, |builder|
        engine::build_capture_program(&*GRAMMAR, "program", source, options, builder));
    fn address(value: &Expression) -> usize {
        let ExprKind::Array(values) = &value.kind else { panic!("array lease"); };
        values.as_ptr() as usize
    }
    with_cache_scope(|| {
        let list = lists(build(MatchOptions::default()).unwrap().into_typed_pairs()).remove(0);
        let original = with_initializer_list(&list, |value| {
            let original = address(value);
            with_cache_scope(|| assert!(take_initializer_list(&list).is_none()));
            assert!(take_initializer_list(&list).is_some(), "reentrant take must not borrow-conflict");
            original
        }).unwrap();
        assert_eq!(address(&take_initializer_list(&list).unwrap()), original,
            "readonly predeclaration must leave the original owned AST for the walk");

        build(MatchOptions::default()).unwrap();
        let mut before_panic = None;
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            with_initializer_list(&list, |value| {
                before_panic = Some(address(value));
                panic!("readonly fold unwinds");
            });
        }));
        assert!(result.is_err());
        assert_eq!(address(&take_initializer_list(&list).unwrap()), before_panic.unwrap());

        build(MatchOptions::default()).unwrap();
        with_initializer_list(&list, |_| {
            assert!(build(MatchOptions { max_steps: 1, ..Default::default() }).is_err());
        }).unwrap();
        assert!(take_initializer_list(&list).is_none(), "lease resurrected a failed parse");

        build(MatchOptions::default()).unwrap();
        let old = with_initializer_list(&list, |value| {
            build(MatchOptions::default()).unwrap();
            address(value)
        }).unwrap();
        assert_ne!(address(&take_initializer_list(&list).unwrap()), old,
            "old lease overwrote a successful same-input reparse");
    });
}

#[test]
fn direct_initializer_replay_observer_counts_constructor_work_not_handoffs() {
    thread_local! {
        static WORK: std::cell::Cell<ReplayWork> = const { std::cell::Cell::new(
            ReplayWork { literals: 0, prefixes: 0, arrays: 0 }) };
    }
    let source = "int a[3]={-1,+2,~3};";
    let build = || build_captures(source, true, |builder|
        engine::build_capture_program(&*GRAMMAR, "program", source, MatchOptions::default(), builder));
    with_replay_observer(|work| WORK.with(|slot| {
        let old = slot.get();
        slot.set(ReplayWork { literals: old.literals + work.literals,
            prefixes: old.prefixes + work.prefixes, arrays: old.arrays + work.arrays });
    }), || with_cache_scope(|| {
        let list = lists(build().unwrap().into_typed_pairs()).remove(0);
        with_initializer_list(&list, |_| ()).unwrap();
        take_initializer_list(&list).unwrap();
        assert_eq!(WORK.with(|slot| slot.get()), ReplayWork::default(),
            "lease and first move must not reconstruct literals");
        take_initializer_list(&list).unwrap();
        assert_eq!(WORK.with(|slot| slot.get()), ReplayWork { literals: 3, prefixes: 3, arrays: 1 });
        take_initializer_list(&list).unwrap();
        assert_eq!(WORK.with(|slot| slot.get()), ReplayWork { literals: 6, prefixes: 6, arrays: 2 });
    }));
    with_cache_scope(|| {
        let list = lists(build().unwrap().into_typed_pairs()).remove(0);
        take_initializer_list(&list).unwrap();
        take_initializer_list(&list).unwrap();
    });
    assert_eq!(WORK.with(|slot| slot.get()), ReplayWork { literals: 6, prefixes: 6, arrays: 2 },
        "observer must not escape its optional profile scope");
}

#[test]
fn direct_initializer_sparse_undo_restores_nested_ownership_anchors() {
    let source = "{{1},2,{3}}";
    let mut builder = Direct::new(source);
    builder.begin(Rule::initializer_list as usize, 0);
    let specifications = [
        (Some(Span { start: 1, end: 4 }), Span { start: 2, end: 3 }),
        (None, Span { start: 5, end: 6 }),
        (Some(Span { start: 7, end: 10 }), Span { start: 8, end: 9 }),
    ];
    let mut expected = Vec::new();
    for (anchor, span) in specifications {
        let leaf = Recipe::Literal { rule: Rule::int_literal, span };
        let recipe = if anchor.is_some() { Recipe::Array(vec![leaf]) } else { leaf };
        let value = recipe.replay(source);
        expected.push(format!("{value:?}"));
        builder.push(Node::Ready { origin: Origin { anchor }, value, recipe });
    }
    let checkpoint = builder.checkpoint();
    builder.finish(Rule::initializer_list as usize,
        Span { start: 0, end: source.len() }, source);
    assert_eq!(builder.depth, 0);
    builder.rollback(checkpoint);
    assert_eq!(builder.depth, 1);
    let values = builder.arena.values();
    assert_eq!(values.len(), 3);
    for (index, id) in values.iter().enumerate() {
        let Node::Ready { origin, value, recipe } = builder.arena.get(*id) else { panic!("restored child"); };
        assert_eq!(origin.anchor, specifications[index].0);
        assert_eq!(format!("{value:?}"), expected[index]);
        assert_eq!(format!("{:?}", recipe.replay(source)), expected[index]);
    }
}

#[test]
fn direct_initializer_rollback_and_budget_errors_match_capture_reference() {
    for source in ["int a[]={1,2,3};", "int a[]={-1,~(-2),+3}; int f( {",
        "int a[]={1,2}; int f( {", "int a[]={{1,},,2};"] {
        for max_steps in [1, 50, 1000, usize::MAX] {
            let mut options = MatchOptions::default();
            options.max_steps = max_steps;
            let control = engine::parse_program(&*GRAMMAR, "program", source, options.clone());
            let mut builder = Direct::with_capture_elision(source);
            let actual = engine::build_capture_program(
                &*GRAMMAR, "program", source, options, &mut builder);
            match (control, actual) {
                (Err(expected), Err(actual)) => {
                    assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                    assert_eq!(builder.into_initializers().len(), 0,
                        "failed parent published a completed child");
                }
                (Ok(_), Ok(_)) => assert!(builder.into_initializers().len() > 0),
                _ => panic!("capture/direct acceptance diverged: {source}, budget={max_steps}"),
            }
        }
    }
}

#[test]
fn direct_initializer_nested_island_survives_contextual_outer_fallback() {
    let source = "int a[2][2]={{1,2},{f(),3}};";
    let (pairs, mut values) = match_program(source, true, MatchOptions::default()).unwrap();
    let lists = lists(pairs);
    assert!(values.take(&lists[0]).is_none());
    let nested = lists.iter().find(|pair| pair.as_str() == "{1,2}").unwrap();
    assert!(values.take(nested).is_some());
    assert_eq!(nested.clone().into_inner().count(), 0);
}

#[test]
fn direct_initializer_owned_nodes_match_complete_module_and_execute_standard_wasm() {
    use vybe_ast::{BindingPattern, StmtKind};
    let source = "int arr[]={-11,+22,~(-34)}; int main(void){return arr[0]+arr[1]+arr[2]==44 ? 0 : 1;}";
    let baseline = crate::parse_with_peg(source).unwrap();
    let mut module = baseline.clone();
    let (pairs, mut values) = match_program(source, true, MatchOptions::default()).unwrap();
    let value = values.take(&lists(pairs).remove(0)).unwrap();
    let declaration = module.body.iter_mut().find_map(|statement| {
        let StmtKind::VarDecl { declarations, .. } = &mut statement.kind else { return None; };
        declarations.iter_mut().find(|declaration|
            matches!(&declaration.pattern, BindingPattern::Ident(name) if name == "arr"))
    }).unwrap();
    declaration.init = Some(value);
    assert_eq!(format!("{module:?}"), format!("{baseline:?}"),
        "direct nodes must preserve full module spans and directives");

    crate::register();
    let mut binaries = Vec::new();
    for module in [baseline, module] {
        let profile = vybe_compiler::profile::parse_profile(crate::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module).unwrap();
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false, include_name_section: true,
            });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap();
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        vm.run(chunks).unwrap();
        assert_eq!(vm.pending_exit_code, 0);
        binaries.push(bytes);
    }
    assert_eq!(binaries[0], binaries[1]);
}
