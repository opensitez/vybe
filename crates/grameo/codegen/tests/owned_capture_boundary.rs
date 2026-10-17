use grameo::{program::Program, CompileOptions};
use grameo_codegen::{emit_with_api, ParserApi};
const BODY:&str="WHITESPACE=_{\" \"} root={SOI~parent~EOI} parent={(owned~\"!\")|(owned~(\"+\"~owned)*)} owned={name~suffix*} name={ASCII_ALPHA+} suffix={\".\"~name|\"(\"~args?~\")\"} args={name~(\",\"~name)*} unrelated={\"z\"}";
fn opted(s: &str) -> grameo::CompiledGrammar {
    grameo::compile_with_options(
        s,
        CompileOptions {
            optimizations: true,
            ..Default::default()
        },
    )
    .unwrap()
}
#[test]
fn selection_default_codegen_and_pratt_mutation_guards() {
    let plain = grameo::compile(BODY).unwrap();
    let source = format!("// @grameo owned_capture_boundary owned\n{BODY}");
    let ignored = grameo::compile(&source).unwrap();
    let base = emit_with_api(&plain, ParserApi::HYBRID_CAPTURES).unwrap();
    assert_eq!(
        base,
        emit_with_api(&ignored, ParserApi::HYBRID_CAPTURES).unwrap()
    );
    let selected = opted(&source);
    let r = selected.rule_id("owned").unwrap();
    let c = selected.owned_capture_boundary(r).unwrap();
    assert_eq!(c.expression, selected.rule(r).expression);
    assert_eq!(c.scope, selected.rule(r).scope);
    let names: Vec<_> = c
        .capture_rules
        .iter()
        .map(|r| selected.rule(*r).name)
        .collect();
    assert_eq!(names, vec!["WHITESPACE", "owned", "name", "suffix", "args"]);
    let code = emit_with_api(&selected, ParserApi::HYBRID_CAPTURES).unwrap();
    assert_eq!(code.matches("fn match_observers_are_noop").count(), 1);
    assert!(code.contains("fn owned_capture_boundary_metadata_is_current"));
    for directive in ["", "missing", "owned owned", "WHITESPACE"] {
        let s = format!("// @grameo owned_capture_boundary {directive}\n{BODY}");
        assert!(grameo::compile(&s).is_ok());
        assert!(grameo::compile_with_options(
            &s,
            CompileOptions {
                optimizations: true,
                ..Default::default()
            }
        )
        .is_err());
    }
    for mode in ["_", "@", "$", "!"] {
        let s = format!("// @grameo owned_capture_boundary owned\nowned={mode}{{\"x\"}}");
        assert!(grameo::compile(&s).is_ok());
        assert!(grameo::compile_with_options(
            &s,
            CompileOptions {
                optimizations: true,
                ..Default::default()
            }
        )
        .is_err());
    }
    let pratt="// @grameo owned_capture_boundary expr\n// @grameo pratt expr atom=atom trailing=consume\n// @grameo operator expr plus precedence=10 fixity=infix_left\nexpr={atom~(plus~atom)*} atom={\"a\"} plus={\"+\"}";
    assert!(grameo::compile_with_options(
        pratt,
        CompileOptions {
            optimizations: true,
            ..Default::default()
        }
    )
    .is_err());
    let child = pratt.replace("boundary expr", "boundary root") + " root={expr}";
    let g = opted(&child);
    let c = g
        .owned_capture_boundary(g.rule_id("root").unwrap())
        .unwrap();
    for n in ["expr", "atom", "plus", "root"] {
        assert!(c.capture_rules.contains(&g.rule_id(n).unwrap()));
    }
    let mut changed = opted(&child);
    assert!(changed.owned_capture_boundary_metadata_is_current());
    changed.set_pratt_rhs_atom("expr", "atom").unwrap();
    assert!(!changed.owned_capture_boundary_metadata_is_current());
    assert!(changed
        .owned_capture_boundary(changed.rule_id("root").unwrap())
        .is_none());
    let mut rebound = opted(&child);
    rebound
        .bind_pratt(
            "expr",
            "atom",
            &[grameo::islands::Operator {
                rule: "plus",
                precedence: 10,
                fixity: grameo::pratt::Fixity::InfixLeft,
            }],
            grameo::islands::TrailingTrivia::Consume,
        )
        .unwrap();
    assert!(!rebound.owned_capture_boundary_metadata_is_current());
    for method in 0..5 {
        let configured = child.replace(
            "trailing=consume",
            "trailing=consume rhs_atom=atom rhs_operator=plus",
        ) + " gap={\"\"}";
        let mut g = opted(&configured);
        assert!(g.owned_capture_boundary_metadata_is_current());
        match method {
            0 => g.set_pratt_operator_trivia("expr", "gap"),
            1 => g.set_pratt_rhs_atom("expr", "atom"),
            2 => g.set_pratt_rhs_operator("expr", "plus"),
            3 => g.set_pratt_operator_rhs_atom("expr", "plus", "atom"),
            _ => g.bind_pratt(
                "expr",
                "atom",
                &[grameo::islands::Operator {
                    rule: "plus",
                    precedence: 10,
                    fixity: grameo::pratt::Fixity::InfixLeft,
                }],
                grameo::islands::TrailingTrivia::Consume,
            ),
        }
        .unwrap();
        assert!(!g.owned_capture_boundary_metadata_is_current());
    }
    println!("selection/default/codegen gates passed; exact scoped call closure excludes unrelated rules");
}
