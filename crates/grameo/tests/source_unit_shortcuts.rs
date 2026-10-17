use grameo::{builder::Builder,engine,source_units::{Collector,WithBuilder,opaque_rules_preserving_requests},CompileOptions,MatchOptions,Span};
#[path = "../../../languages/php/src/walker/source_unit_admission.rs"]
mod php_source_unit_admission;

struct LexicalOwner;
impl Builder for LexicalOwner {
    fn checkpoint(&self)->usize {0}
    fn rollback(&mut self,_:usize) {}
    fn observes_rule(&self,_:usize)->bool {false}
    fn opaque_rule(&self,_:usize)->bool {true}
    fn supports_shared_group_prefix(&self,_:usize,_:usize)->bool {true}
    fn begin(&mut self,_:usize,_:usize) {panic!("unobserved begin");}
    fn finish(&mut self,_:usize,_:Span,_:&str) {panic!("unobserved finish");}
}

#[test]
fn actual_php_lexical_shortcuts_keep_all_committed_load_requests() {
    let source=include_str!("../../../languages/php/src/grammar.pest");
    let grammar=grameo::compile_with_options(source,CompileOptions {optimizations:true,..Default::default()}).unwrap();
    let id=|name|grammar.rule_id(name).unwrap();
    let bindings=grammar.source_unit_bindings().collect::<Vec<_>>();
    let admitted=opaque_rules_preserving_requests(&grammar,&[id("program"),id("program_pure")],&bindings);
    for name in ["kw_include","kw_include_once","kw_require","kw_require_once","number_lit","sq_string","dq_string","simple_variable"] {
        assert!(admitted.contains(&id(name)),"{name}");
    }
    assert!(!admitted.contains(&id("include_expression")));
    for input in [
        "include 'a.php'; include_once ('b.php'); require 'c.php'; require_once $filename;",
        "function load($basedir) { include $basedir . '/x.php'; }",
        "include 'first.php'; function f() { include 'function.php'; } class C { function m() { require 'method.php'; } } $f = function () { include 'closure.php'; }; $g = fn() => require_once 'arrow.php'; include 'last.php';",
        "include (include 'path.php');",
        "if (1 + 2 > 0) { require_once 'inside.php'; }",
        "$string = 'include fake.php'; /* require fake.php */ include 'real.php';",
    ] {
        let mut control=Collector::new(bindings.iter().copied());
        control.defer_within(["function_declaration","method_declaration","closure_expression","arrow_function"].map(id));
        engine::build_capture_program(&grammar,"program_pure",input,MatchOptions::default(),&mut control).unwrap();
        let mut combined=WithBuilder::new(LexicalOwner,bindings.iter().copied());
        combined.collector.defer_within(["function_declaration","method_declaration","closure_expression","arrow_function"].map(id));
        combined.allow_opaque_rules(admitted.iter().copied());
        engine::build_capture_program(&grammar,"program_pure",input,MatchOptions::default(),&mut combined).unwrap();
        assert_eq!(combined.collector.owned_requests(),control.owned_requests(),"{input}");
        assert!(!combined.collector.owned_requests().is_empty());
    }
}

#[test]
fn source_guard_cannot_override_owner_pratt_refusal() {
    let grammar=grameo::compile_with_options("// @grameo source_unit load target=name loader=source phase=execute repeat=always missing=error scope=caller\nroot={load~EOI} load={\"[\"~name~\"]\"} name=@{ASCII_ALPHA+}",CompileOptions {optimizations:true,..Default::default()}).unwrap();
    let mut combined=WithBuilder::new(LexicalOwner,grammar.source_unit_bindings());
    combined.set_pratt_admission(|_,_,_,_|true);
    assert!(!combined.supports_pratt_at(grammar.rule_id("name").unwrap(),0,"file"));
    assert!(!combined.supports_shared_group_prefix(0,1));
    let unbound=WithBuilder::new(LexicalOwner,std::iter::empty());
    assert!(unbound.supports_shared_group_prefix(0,1));
}

#[test]
fn php_negative_load_admission_never_hides_a_grammar_load() {
    let grammar=grameo::compile_with_options(include_str!("../../../languages/php/src/grammar.pest"),
        CompileOptions {optimizations:true,..Default::default()}).unwrap();
    for keyword in ["include","include_once","require","require_once"] {
        // All case variants of the semantic stem, with each suffix spelling.
        // The actual grammar supplies the receipts; this is not a second lexer.
        for mask in 0..128 {
            let mut spelling=keyword.as_bytes().to_vec();
            for (bit,byte) in spelling.iter_mut().take(7).enumerate() {
                if mask & (1 << bit)!=0 {*byte=byte.to_ascii_uppercase();}
            }
            let input=format!("{} 'unit.php';",String::from_utf8(spelling).unwrap());
            let mut collector=Collector::new(grammar.source_unit_bindings());
            engine::build_capture_program(&grammar,"program_pure",&input,MatchOptions::default(),&mut collector).unwrap();
            assert_eq!(collector.requests().len(),1,"{input}");
            assert!(php_source_unit_admission::may_contain_source_load(&input),"{input}");
        }
    }
    for input in ["$x = f($a) + 2;", "$x = '日本語';", "$x = $a ?? $b;", "function f() { return g(2); }"] {
        let mut collector=Collector::new(grammar.source_unit_bindings());
        engine::build_capture_program(&grammar,"program_pure",input,MatchOptions::default(),&mut collector).unwrap();
        assert!(collector.requests().is_empty());
        assert!(!php_source_unit_admission::may_contain_source_load(input));
    }
    // These deliberately keep the slower collection route. No lexical guess
    // about strings, comments, or identifier boundaries can hide a real load.
    for input in ["$x = 'REQUIRE';", "/* include */ $x=1;", "$include_count=1;", "日本語require_once"] {
        assert!(php_source_unit_admission::may_contain_source_load(input));
    }
}
