use grameo::{Builtin, CompiledGrammar, RuleMode, program::{Program, Instruction}};
use std::{collections::BTreeSet, fmt::Write};
#[derive(Default)]
struct Head { words: Vec<(String, bool)>, class: [u64; 4], rules: BTreeSet<usize>, depth: usize }
fn bit(class: &mut [u64;4], byte:u8) { class[usize::from(byte/64)] |= 1u64 << (byte%64); }
fn walk(g:&CompiledGrammar,id:usize,scope:usize,level:usize,head:&mut Head)->Option<()> {
    if level > 96 { return None }
    match g.instruction(id) {
        Instruction::Literal{text,insensitive} if !text.is_empty() && text.is_ascii() => {
            head.words.push((text.into(),insensitive));
        }
        Instruction::Builtin(builtin) => {
            for byte in 0u8..128 {
                let accepts=match builtin {
                    Builtin::Ascii => true,
                    Builtin::AsciiDigit => byte.is_ascii_digit(),
                    Builtin::AsciiAlpha => byte.is_ascii_alphabetic(),
                    Builtin::AsciiAlphaLower => byte.is_ascii_lowercase(),
                    Builtin::AsciiAlphaUpper => byte.is_ascii_uppercase(),
                    Builtin::AsciiAlphanumeric => byte.is_ascii_alphanumeric(),
                    Builtin::AsciiHexDigit => byte.is_ascii_hexdigit(),
                    Builtin::AsciiNonzeroDigit => (b'1'..=b'9').contains(&byte),
                    Builtin::AsciiBinDigit => matches!(byte,b'0'|b'1'),
                    Builtin::AsciiOctDigit => (b'0'..=b'7').contains(&byte),
                    _ => return None,
                };
                if accepts { bit(&mut head.class,byte); }
            }
        }
        Instruction::Range{start,end} => {
            for byte in 0u8..128 { if (start..=end).contains(&char::from(byte)) {bit(&mut head.class,byte)} }
            if end>'\u{7f}' { for byte in 0xc2..=0xf4 {bit(&mut head.class,byte)} }
        }
        Instruction::Call(rule) => {
            let spec=g.rule(rule);
            if spec.scope!=scope || matches!(spec.mode,RuleMode::NonAtomic|RuleMode::CompoundAtomic)
                || g.source_pratt(rule).is_some() {return None}
            let before=head.depth; head.depth=0;
            walk(g,spec.expression,scope,level+1,head)?;
            head.depth=before.max(head.depth+1);
            head.rules.insert(rule);
        }
        Instruction::Group(child) | Instruction::Predicate{expression:child,positive:true} => walk(g,child,scope,level+1,head)?,
        Instruction::Sequence{left,..} => walk(g,left,scope,level+1,head)?,
        Instruction::Choice{left,right} => {walk(g,left,scope,level+1,head)?;walk(g,right,scope,level+1,head)?;}
        Instruction::Repeat{expression,min,..} if min>0 => walk(g,expression,scope,level+1,head)?,
        _ => return None,
    }
    Some(())
}
pub(super) fn emit(out:&mut String,g:&CompiledGrammar) {
    let mut roots=Vec::new();let mut branch_rules=String::new();
    out.push_str("    fn choice_head_branch_metadata_is_current(&self)->bool{true}\n");
    if !out.contains("    fn match_observers_are_noop(") {out.push_str("    fn match_observers_are_noop(&self)->bool{true}\n");}
    out.push_str("    fn choice_heads_metadata_is_current(&self)->bool{true}\n    fn choice_heads(&self,id:usize)->Option<&::grameo::program::ChoiceHeads>{match id {\n");
    for &owner in g.choice_head_rules(){
        let rule=g.rule(owner);if !matches!(rule.mode,RuleMode::Normal|RuleMode::Silent) || g.source_pratt(owner).is_some(){continue;}
        let mut id=rule.expression;while let Instruction::Group(c)=g.instruction(id){id=c;}
        if !matches!(g.instruction(id),Instruction::Choice{..}){continue;}
        let mut branches=Vec::new();fn flatten(g:&CompiledGrammar,id:usize,out:&mut Vec<usize>){match g.instruction(id){Instruction::Choice{left,right}=>{flatten(g,left,out);flatten(g,right,out);},_=>out.push(id)}}
        flatten(g,id,&mut branches);if branches.len()>64{continue;}
        let mut rules=BTreeSet::new();let mut depth=0;let mut emitted=String::new();let mut known=0;let mut first=[0u64;256];
        for (index,branch) in branches.into_iter().enumerate() {
            let mut head=Head::default();write!(emitted,"::grameo::program::ChoiceHeadBranch{{expression:{branch},head:").unwrap();
            if walk(g,branch,rule.scope,0,&mut head).is_some(){
                for byte in 0u8..=255 {
                    if head.class[usize::from(byte/64)]&(1u64<<(byte%64))!=0
                        || head.words.iter().any(|(word,fold)|word.as_bytes()[0]==byte || *fold && word.as_bytes()[0].eq_ignore_ascii_case(&byte)) {first[usize::from(byte)]|=1u64<<index;}
                }
                known+=1;depth=depth.max(head.depth);
                writeln!(branch_rules,"({id},{index})=>Some(&{:?}),",head.rules.iter().copied().collect::<Vec<_>>()).unwrap();
                rules.extend(head.rules);
                write!(emitted,"Some(&::grameo::program::ChoiceHead{{class:{:?},words:&{:?}}})",head.class,head.words).unwrap();
            }else{for lane in &mut first{*lane|=1u64<<index;}emitted.push_str("None");}
            emitted.push_str("},");
        }
        if known==0{continue;}roots.push(id);
        writeln!(out,"{id}=>Some(&::grameo::program::ChoiceHeads{{first:&{first:?},scope:{},expression:{id},branches:&[{emitted}],rules:&{:?},depth:{depth}}}),",rule.scope,rules.into_iter().collect::<Vec<_>>()).unwrap();
    }
    out.push_str("_=>None,}}\n");
    writeln!(out,"    fn choice_head_branch_rules(&self,id:usize,branch:usize)->Option<&[usize]>{{match(id,branch){{{branch_rules}_=>None,}}}}").unwrap();
    if roots.is_empty(){out.push_str("    fn choice_heads_candidate(&self,_id:usize)->bool{false}\n");}
    else{let pattern=roots.iter().map(usize::to_string).collect::<Vec<_>>().join("|");writeln!(out,"    #[inline]\n    fn choice_heads_candidate(&self,id:usize)->bool{{matches!(id,{pattern})}}").unwrap();}
}
