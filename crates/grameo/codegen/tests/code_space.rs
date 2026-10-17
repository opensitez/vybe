use grameo::{compile, compile_with_options, CompileOptions, CompiledGrammar, Diagnostic};
use grameo::program::Program;
use grameo::source_units::{CodeSpace, Missing, Repetition, Scope};

const BODY: &str = "entry = { SOI ~ load ~ EOI }\nload = { (once | include) ~ target }\ninclude = @{ \"include\" }\nonce = @{ \"include_once\" }\ntarget = @{ ASCII_ALPHA+ }\n";

fn optimized(source: &str) -> Result<CompiledGrammar, Vec<Diagnostic>> {
    compile_with_options(source, CompileOptions { optimizations: true, ..Default::default() })
}

#[test]
fn peg_mode_ignores_even_invalid_code_space_comments() {
    let baseline = compile(BODY).unwrap();
    for policy in ["request", "unit", "unknown", "request unit", ""] {
        let grammar = compile(&format!("// @grameo code_space {policy}\n{BODY}")).unwrap();
        assert_eq!(grammar.code_space(), CodeSpace::Unit);
        assert_eq!(grameo_codegen::emit(&grammar).unwrap(), grameo_codegen::emit(&baseline).unwrap());
        for input in ["includeabc", "include_onceabc", "broken"] {
            assert_eq!(grammar.recognize("entry", input).map(|r| r.consumed),
                       baseline.recognize("entry", input).map(|r| r.consumed));
        }
    }
}

#[test]
fn request_storage_requires_explicit_opt_in_and_is_emitted() {
    assert_eq!(optimized(BODY).unwrap().code_space(), CodeSpace::Unit);
    for (policy, expected) in [("unit", CodeSpace::Unit), ("request", CodeSpace::Request)] {
        let grammar = optimized(&format!("// @grameo code_space {policy}\n{BODY}")).unwrap();
        assert_eq!(grammar.code_space(), expected);
        assert!(grammar.recognize("entry", "includeabc").is_ok());
        assert_eq!(grameo_codegen::emit(&grammar).unwrap().contains("fn code_space("),
                   expected == CodeSpace::Request);
    }
}

#[test]
fn optimized_mode_rejects_invalid_or_duplicate_storage_contracts() {
    for policy in ["", "unknown", "request unit", "request extra=true"] {
        assert!(optimized(&format!("// @grameo code_space {policy}\n{BODY}")).is_err());
    }
    for first in ["unit", "request"] {
        for second in ["unit", "request"] {
            assert!(optimized(&format!("// @grameo code_space {first}\n// @grameo code_space {second}\n{BODY}")).is_err());
        }
    }
}

#[test]
fn storage_contract_preserves_include_scope_repeat_and_failure_policy() {
    let declarations = "// @grameo source_unit load target=target selector=include loader=source phase=execute repeat=always missing=warning scope=caller\n// @grameo source_unit load target=target selector=once loader=source phase=execute repeat=once missing=error scope=caller\n";
    let baseline = optimized(&format!("{declarations}{BODY}")).unwrap();
    let grammar = optimized(&format!("// @grameo code_space request\n{declarations}{BODY}")).unwrap();
    for selector in ["include", "once"] {
        let before = baseline.source_unit_binding(baseline.rule_id("load").unwrap(), baseline.rule_id(selector)).unwrap();
        let after = grammar.source_unit_binding(grammar.rule_id("load").unwrap(), grammar.rule_id(selector)).unwrap();
        assert_eq!(before, after);
        assert_eq!(after.policy.scope, Scope::Caller);
        if selector == "once" {
            assert_eq!(after.policy.repetition, Repetition::Once);
            assert_eq!(after.policy.missing, Missing::Error);
        }
    }
}
