use grameo::{
    MatchOptions, Pair, ParseTree,
    engine::{parse_program},
    program::{Instruction, Program, RuleSpec, Scope},
};

#[derive(Debug)]
struct Scheduled;
use vybe_parser_generated_tests::terminal_index as generated;
impl Program for Scheduled {
 fn atomic_expression(&self,id:usize)->bool { generated::Parser.atomic_expression(id) }
 fn leading_bytes(&self,id:usize)->Option<grameo::lexical::LeadingBytes> { generated::Parser.leading_bytes(id) }
 fn terminal_choice(&self,id:usize)->Option<grameo::lexical::TerminalChoice<'_>> { generated::Parser.terminal_choice(id) }
 fn failure_terminals(&self,id:usize)->&[usize] { generated::Parser.failure_terminals(id) }
 fn literal_choice_failure(&self,id:usize,input:&str)->Option<usize> { generated::Parser.literal_choice_failure(id,input) }
 fn trivia_failure_prefix(&self,id:usize)->Option<grameo::lexical::FailurePrefix> { generated::Parser.trivia_failure_prefix(id) }
 fn fast_ascii_repetitions(&self)->bool { generated::Parser.fast_ascii_repetitions() }
 fn rule_id(&self,name:&str)->Option<usize> { generated::Parser.rule_id(name) }
 fn rule(&self,id:usize)->RuleSpec<'_> { generated::Parser.rule(id) }
 fn instruction(&self,id:usize)->Instruction<'_> { generated::Parser.instruction(id) }
 fn scope(&self,id:usize)->Scope { generated::Parser.scope(id) }
}

fn shape(
    tree: ParseTree<'_, '_>,
) -> (
    usize,
    Vec<(grameo::CaptureRule, grameo::Span, usize)>,
) {
    fn visit(
        pair: Pair<'_, '_, '_>,
        depth: usize,
        out: &mut Vec<(grameo::CaptureRule, grameo::Span, usize)>,
    ) {
        out.push((pair.as_rule(), pair.as_span(), depth));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut out = Vec::new();
    for pair in tree.pairs() {
        visit(pair, 0, &mut out);
    }
    assert_eq!(out.len(), tree.capture_count());
    (tree.consumed(), out)
}
#[test]
fn indexed_literal_choices_preserve_every_short_resource_boundary() {
 for entry in ["root","positive","visible","speculate"] {
 for input in ["abc","abcz","","if","ifdef","IF","ifx","while","wh","for","local","true","false","nil","other","é","éx","中","😀"] {
 for depth in 0..=4 {
 for budget in 0..=150 {
 let options=MatchOptions{max_steps:budget,max_rule_depth:depth};
 assert_eq!(parse_program(&generated::Parser,entry,input,options).map(shape),parse_program(&Scheduled,entry,input,options).map(shape),"{entry} {input:?} {options:?}");
 }}}}
}
