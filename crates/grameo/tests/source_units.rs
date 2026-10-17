use grameo::{CompileOptions, compile, compile_with_options, program::Program};
use grameo::source_units::{Phase, Repetition, Missing, Scope};

const GRAMMAR: &str = r#"
load = { (include | once | require | require_once) ~ target }
include = @{ "include" }
once = @{ "include_once" }
require = @{ "require" }
require_once = @{ "require_once" }
target = @{ ASCII_ALPHA+ }
unrelated = { "other" }
"#;
fn opted(source: &str) -> Result<grameo::CompiledGrammar, Vec<grameo::Diagnostic>> {
    compile_with_options(source, CompileOptions {optimizations:true, ..Default::default()})
}
#[test]
fn default_peg_ignores_even_malformed_source_unit_directives() {
    let source = format!("// @grameo source_unit nonsense garbage\n{GRAMMAR}");
    let g = compile(&source).unwrap();
    assert_eq!(g.source_unit_bindings().count(),0);
    assert!(g.source_unit_binding(g.rule_id("load").unwrap(), None).is_none());
    assert!(opted(&source).is_err());
}
#[test]
fn four_forms_keep_independent_repetition_and_failure_policies() {
    let mut source = String::new();
    for (selector, repeat, missing) in [("include","always","warning"), ("once","once","warning"), ("require","always","error"), ("require_once","once","error")] {
        source.push_str(&format!("// @grameo source_unit load target=target selector={selector} loader=source phase=execute repeat={repeat} missing={missing} scope=caller\n"));
    }
    source.push_str(GRAMMAR);
    let g = opted(&source).unwrap();
    assert_eq!(g.source_unit_bindings().count(),4);
    let root = g.rule_id("load").unwrap();
    for (selector, repeat, missing) in [("include",Repetition::Always,Missing::Warning), ("once",Repetition::Once,Missing::Warning), ("require",Repetition::Always,Missing::Error), ("require_once",Repetition::Once,Missing::Error)] {
        let b = g.source_unit_binding(root, g.rule_id(selector)).unwrap();
        assert_eq!(b.policy.phase,Phase::Execute);
        assert_eq!(b.policy.scope,Scope::Caller);
        assert_eq!(b.policy.repetition,repeat);
        assert_eq!(b.policy.missing,missing);
        assert_eq!(b.target,g.rule_id("target").unwrap());
        let request = b.request("unevaluated_target", grameo::Span {start:4,end:20});
        assert_eq!(request.target,"unevaluated_target");
        assert_eq!(request.policy,b.policy);
        assert_eq!(request.span.start,4);
    }
    assert!(g.source_unit_binding(root,None).is_none());
}
#[test]
fn bindings_reject_unknown_duplicate_or_unreachable_properties() {
    let good="load target=target loader=source phase=module repeat=once missing=error scope=module";
    for bad in [good.replace("target=target","target=absent"), good.replace("target=target","target=unrelated"), good.replace("scope=module","scope=caller scope=module"), good.replace("phase=module","phase=guess"), format!("{good} typo=true")] {
        assert!(opted(&format!("// @grameo source_unit {bad}\n{GRAMMAR}")).is_err(),"{bad}");
    }
    assert!(opted(&format!("// @grameo source_unit {good}\n// @grameo source_unit {good}\n{GRAMMAR}")).is_err());
}
#[test]
fn predicate_only_operand_is_not_semantic_source_intent() {
    let source="// @grameo source_unit load target=target loader=source phase=execute repeat=always missing=error scope=explicit\nload = { &target ~ ANY }\ntarget = { ANY }";
    assert!(opted(source).is_err());
}

#[test]
fn repository_php_declares_all_four_forms_without_changing_default_peg() {
    let source = include_str!("../../../languages/php/src/grammar.pest");
    let default = compile(source).unwrap();
    assert_eq!(default.source_unit_bindings().count(),0);
    let grammar = opted(source).unwrap();
    let root = grammar.rule_id("include_expression").unwrap();
    for (keyword,repeat,missing) in [("kw_include",Repetition::Always,Missing::Warning),("kw_include_once",Repetition::Once,Missing::Warning),("kw_require",Repetition::Always,Missing::Error),("kw_require_once",Repetition::Once,Missing::Error)] {
        let binding = grammar.source_unit_binding(root,grammar.rule_id(keyword)).unwrap();
        assert_eq!(binding.target,grammar.rule_id("include_operand").unwrap());
        assert_eq!(binding.policy.repetition,repeat);
        assert_eq!(binding.policy.missing,missing);
        assert_eq!(binding.policy.scope,Scope::Caller);
    }
}

#[test]
fn collector_commits_nested_requests_and_rolls_back_speculation_and_failures() {
    use grameo::{engine,MatchOptions,source_units::Collector};
    let source=r#"
// @grameo source_unit load target=target selector=include loader=source phase=execute repeat=always missing=warning scope=caller
// @grameo source_unit load target=target selector=once loader=source phase=execute repeat=once missing=error scope=explicit
root = { &load ~ (load ~ "!" | load) ~ EOI }
load = { (include | once) ~ "(" ~ target ~ ")" }
include = @{ "include" }
once = @{ "once" }
target = { load | variable }
variable = @{ "$" ~ ASCII_ALPHA+ }
"#;
    let grammar=opted(source).unwrap();
    let input="once(include($basedir))";
    for captures in [false,true] {
        let mut collector=Collector::new(grammar.source_unit_bindings());
        if captures {
            engine::build_capture_program(&grammar,"root",input,MatchOptions::default(),&mut collector).unwrap();
        } else {
            engine::build_ast_program(&grammar,"root",input,MatchOptions::default(),&mut collector).unwrap();
        }
        assert_eq!(collector.requests().len(),2);
        assert_eq!(&input[collector.requests()[0].target.start..collector.requests()[0].target.end],"$basedir");
        assert_eq!(&input[collector.requests()[1].target.start..collector.requests()[1].target.end],"include($basedir)");
        assert_eq!(collector.requests()[0].policy.scope,Scope::Caller);
        assert_eq!(collector.requests()[1].policy.scope,Scope::Explicit);
        let before=collector.requests().to_vec();
        assert!(engine::build_capture_program(&grammar,"root","once(include($basedir))?",MatchOptions::default(),&mut collector).is_err());
        assert_eq!(collector.requests(),before);
        assert_eq!(collector.take_requests(),before);
        assert!(collector.requests().is_empty());
    }
}

#[test]
fn collector_handles_target_before_selector_and_rejects_ambiguous_targets() {
    use grameo::{engine,MatchOptions,source_units::Collector};
    for (body,input,valid) in [("target ~ keyword","fileload",true),("target ~ target ~ keyword","filefileload",false)] {
        let source=format!("// @grameo source_unit load target=target selector=keyword loader=source phase=module repeat=once missing=diagnostic scope=module\nload = {{ {body} ~ EOI }} target = @{{ \"file\" }} keyword = @{{ \"load\" }}");
        let grammar=opted(&source).unwrap();
        let mut collector=Collector::new(grammar.source_unit_bindings());
        let result=engine::build_ast_program(&grammar,"load",input,MatchOptions::default(),&mut collector);
        assert_eq!(result.is_ok(),valid);
        assert_eq!(collector.requests().len(),usize::from(valid));
        if valid {assert_eq!(collector.requests()[0].policy.phase,Phase::Module);}
    }
}

#[test]
fn php_collector_keeps_computed_filename_unevaluated_in_function_scope() {
    use grameo::{engine,MatchOptions,source_units::Collector};
    let grammar=opted(include_str!("../../../languages/php/src/grammar.pest")).unwrap();
    let input="function f($basedir) { return include_once ($basedir . '/file.php'); }";
    let mut collector=Collector::new(grammar.source_unit_bindings());
    engine::build_capture_program(&grammar,"program_pure",input,MatchOptions::default(),&mut collector).unwrap();
    assert_eq!(collector.requests().len(),1);
    let request=&collector.requests()[0];
    assert_eq!(input[request.target.start..request.target.end].trim(),"($basedir . '/file.php')");
    assert_eq!(request.policy.scope,Scope::Caller);
    assert_eq!(request.policy.repetition,Repetition::Once);
    assert_eq!(collector.owned_requests()[0].target,"($basedir . '/file.php')");
}

#[test]
fn owned_receipts_retain_exact_normalized_input_across_reuse_and_failure() {
    use grameo::{engine,MatchOptions,source_units::Collector};
    let grammar=opted("// @grameo source_unit load target=target loader=source phase=execute repeat=always missing=error scope=caller\nload = { \"[\" ~ target ~ \"]\" ~ EOI } target = @{ ASCII_ALPHA+ }").unwrap();
    let mut collector=Collector::new(grammar.source_unit_bindings());
    let mut normalized=String::from("[first]");
    engine::build_capture_program(&grammar,"load",&normalized,MatchOptions::default(),&mut collector).unwrap();
    normalized.clear();normalized.push_str("[other]");
    assert_eq!(collector.owned_requests()[0].target,"first");
    assert!(engine::build_capture_program(&grammar,"load","[failed]?",MatchOptions::default(),&mut collector).is_err());
    assert_eq!(collector.owned_requests().len(),1);
    engine::build_capture_program(&grammar,"load",&normalized,MatchOptions::default(),&mut collector).unwrap();
    let requests=collector.take_owned_requests();
    assert_eq!(requests.iter().map(|request|request.target.as_str()).collect::<Vec<_>>(),["first","other"]);
    assert!(collector.owned_requests().is_empty());
}

#[test]
fn collector_resource_limits_restore_all_pending_requests() {
    use grameo::{engine,MatchOptions,source_units::Collector};
    let source="// @grameo source_unit load target=target loader=source phase=preprocess repeat=always missing=diagnostic scope=global\nroot = { load+ ~ EOI } load = { \"[\" ~ target ~ \"]\" } target = @{ ASCII_ALPHA+ }";
    let grammar=opted(source).unwrap();
    let input="[one][two][three]";
    for steps in 0..120 {
        for depth in [1,2,3,32] {
            let mut collector=Collector::new(grammar.source_unit_bindings());
            let result=engine::build_capture_program(&grammar,"root",input,MatchOptions {max_steps:steps,max_rule_depth:depth},&mut collector);
            if result.is_err() {
                assert!(collector.requests().is_empty());
                // The same collector must be usable after work/depth failure.
                engine::build_capture_program(&grammar,"root",input,MatchOptions::default(),&mut collector).unwrap();
            }
            assert_eq!(collector.requests().len(),3);
            assert_eq!(collector.requests()[0].policy.phase,Phase::Preprocess);
            assert_eq!(collector.requests()[0].policy.scope,Scope::Global);
        }
    }
}

#[test]
fn unregistered_collector_preserves_default_peg_and_publishes_nothing() {
    use grameo::{engine,MatchOptions,source_units::Collector};
    let grammar=compile(&format!("// @grameo source_unit invalid garbage\n{GRAMMAR}")).unwrap();
    let mut collector=Collector::new(grammar.source_unit_bindings());
    engine::build_capture_program(&grammar,"load","includefile",MatchOptions::default(),&mut collector).unwrap();
    assert!(collector.requests().is_empty());
}

#[test]
fn nested_source_unit_can_be_another_units_direct_target() {
    use grameo::{engine,MatchOptions,source_units::Collector};
    let grammar=opted("// @grameo source_unit outer target=inner loader=source phase=execute repeat=always missing=error scope=caller\n// @grameo source_unit inner target=name loader=source phase=module repeat=once missing=error scope=module\nroot = { (outer ~ \"!\" | outer) ~ EOI } outer = { \"[\" ~ inner ~ \"]\" } inner = { \"(\" ~ name ~ \")\" } name = @{ ASCII_ALPHA+ }").unwrap();
    let mut collector=Collector::new(grammar.source_unit_bindings());
    engine::build_ast_program(&grammar,"root","[(file)]",MatchOptions::default(),&mut collector).unwrap();
    assert_eq!(collector.requests().len(),2);
    assert_eq!(collector.requests()[0].target,grameo::Span {start:2,end:6});
    assert_eq!(collector.requests()[1].target,grameo::Span {start:1,end:7});
}
