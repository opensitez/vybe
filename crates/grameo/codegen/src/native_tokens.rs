//! Opt-in direct lexical recognizers. Pure PEG success paths only; failures replay.
use grameo::{
    Builtin, CompiledGrammar, RuleMode,
    program::{Instruction, Program},
};
use std::{collections::BTreeSet, fmt::Write};
fn ascii(b: Builtin) -> Option<&'static str> {
    Some(match b {
        Builtin::Ascii => "b<128",
        Builtin::AsciiDigit => "b.is_ascii_digit()",
        Builtin::AsciiAlpha => "b.is_ascii_alphabetic()",
        Builtin::AsciiAlphaLower => "b.is_ascii_lowercase()",
        Builtin::AsciiAlphaUpper => "b.is_ascii_uppercase()",
        Builtin::AsciiAlphanumeric => "b.is_ascii_alphanumeric()",
        Builtin::AsciiHexDigit => "b.is_ascii_hexdigit()",
        Builtin::AsciiNonzeroDigit => "(b'1'..=b'9').contains(&b)",
        Builtin::AsciiBinDigit => "b==b'0'||b==b'1'",
        Builtin::AsciiOctDigit => "(b'0'..=b'7').contains(&b)",
        _ => return None,
    })
}
fn certify(
    g: &CompiledGrammar,
    id: usize,
    path: &mut BTreeSet<usize>,
    used: &mut BTreeSet<usize>,
    level: usize,
) -> Option<usize> {
    if level > 128 || !path.insert(id) {
        return None;
    }
    let mut child = |id| certify(g, id, path, used, level + 1);
    let depth = match g.instruction(id) {
        Instruction::Literal { text, .. } if text.is_ascii() => Some(0),
        Instruction::Range { start, end } if start.is_ascii() && end.is_ascii() => Some(0),
        Instruction::Builtin(b)
            if ascii(b).is_some()
                || matches!(b, Builtin::Any | Builtin::Soi | Builtin::Eoi | Builtin::Newline) =>
        {
            Some(0)
        }
        Instruction::Sequence { left, right } | Instruction::Choice { left, right } => {
            Some(child(left)?.max(child(right)?))
        }
        Instruction::Group(c) | Instruction::Predicate { expression: c, .. } => child(c),
        Instruction::Repeat { expression: c, .. }
            if !g.analysis().expressions[c].may_succeed_empty =>
        {
            child(c)
        }
        Instruction::Call(r)
            if !matches!(
                g.rule(r).mode,
                RuleMode::NonAtomic | RuleMode::CompoundAtomic
            ) =>
        {
            child(g.rule(r).expression)?.checked_add(1)
        }
        _ => None,
    };
    path.remove(&id);
    if depth.is_some() {
        used.insert(id);
    }
    depth
}
pub(super) fn emit(out: &mut String, g: &CompiledGrammar) {
    if !g.native_token_compilation() {
        return;
    }
    let mut roots = Vec::new();
    let mut used = BTreeSet::new();
    for r in 0..g.syntax().rules.len() {
        let rule = g.rule(r);
        if rule.mode != RuleMode::Atomic || !g.atomic_expression(rule.expression) {
            continue;
        }
        // Prefer an existing proof for a more specialized whole-rule scanner.
        // Children can still be compiled as parts of other certified roots.
        if grameo::quoted::certify(g, r).is_some()
            || grameo::quoted::certify_delimited(g, r).is_some()
        {
            continue;
        }
        let mut candidate = BTreeSet::new();
        if let Some(depth) = certify(g, rule.expression, &mut BTreeSet::new(), &mut candidate, 0) {
            let scalar_boundary = candidate.iter().any(|&id|
                matches!(g.instruction(id), Instruction::Builtin(Builtin::Any)));
            roots.push((r, rule.expression, depth, scalar_boundary));
            used.extend(candidate);
        }
    }
    if roots.is_empty() {
        return;
    }
    out.push_str("    fn native_token_metadata_is_current(&self)->bool{true}\n");
    let failures = g.native_token_failures();
    if failures {
        out.push_str("    fn native_token(&self, rule:usize, source:&str, offset:usize, remaining_depth:usize)->Option<usize>{self.native_token_result(rule,source,offset,remaining_depth).flatten()}\n");
        out.push_str("    fn native_token_result(&self, rule:usize, source:&str, offset:usize, remaining_depth:usize)->Option<Option<usize>>{\n        if !source.is_char_boundary(offset){return None;}\n");
    } else {
        out.push_str("    fn native_token(&self, rule:usize, source:&str, offset:usize, remaining_depth:usize)->Option<usize>{\n");
    }
    for id in used {
        writeln!(
            out,
            "        #[inline] fn t{id}(s:&[u8],p:usize)->Option<usize>{{"
        )
        .unwrap();
        match g.instruction(id){
   Instruction::Literal{text,insensitive}=>{if insensitive{writeln!(out,"s.get(p..p.checked_add({})?).filter(|b|b.eq_ignore_ascii_case({:?}.as_bytes())).map(|_|p+{})",text.len(),text,text.len()).unwrap();}else{writeln!(out,"s.get(p..p.checked_add({})?).filter(|b|*b=={:?}.as_bytes()).map(|_|p+{})",text.len(),text,text.len()).unwrap();}}
   Instruction::Range{start,end}=>{writeln!(out,"s.get(p).filter(|&&b|({}..={}).contains(&b)).map(|_|p+1)",start as u8,end as u8).unwrap();}
   Instruction::Builtin(b)=>{match b{Builtin::Any=>out.push_str("let b=*s.get(p)?;if b<128{return p.checked_add(1)}let width=match b{0xc2..=0xdf=>2,0xe0..=0xef=>3,0xf0..=0xf4=>4,_=>return None};let end=p.checked_add(width)?;std::str::from_utf8(s.get(p..end)?).ok()?;Some(end)"),Builtin::Soi=>out.push_str("(p==0).then_some(p)"),Builtin::Eoi=>out.push_str("(p==s.len()).then_some(p)"),Builtin::Newline=>out.push_str("if s.get(p..p+2)==Some(b\"\\r\\n\"){Some(p+2)}else{s.get(p).filter(|&&b|b==b'\\n'||b==b'\\r').map(|_|p+1)}"),_=>{writeln!(out,"s.get(p).filter(|&&b|{}).map(|_|p+1)",ascii(b).unwrap()).unwrap();}}}
   Instruction::Call(r)=>{writeln!(out,"t{}(s,p)",g.rule(r).expression).unwrap();}
   Instruction::Group(c)=>{writeln!(out,"t{c}(s,p)").unwrap();}
   Instruction::Sequence{left,right}=>{writeln!(out,"let q=t{left}(s,p)?;t{right}(s,q)").unwrap();}
   Instruction::Choice{left,right}=>{writeln!(out,"t{left}(s,p).or_else(||t{right}(s,p))").unwrap();}
   Instruction::Predicate{expression,positive}=>{writeln!(out,"(t{expression}(s,p).is_some()=={positive}).then_some(p)").unwrap();}
   Instruction::Repeat{expression,min,max}=>{writeln!(out,"let mut q=p;let mut n=0u32;while n<{}{{match t{expression}(s,q){{Some(end)=>{{q=end;n+=1;}},None=>break}}}}(n>={min}).then_some(q)",max.unwrap_or(u32::MAX)).unwrap();}
   _=>unreachable!(),
  }
        out.push_str("        }\n");
    }
    out.push_str("        match rule{\n");
    for (r, id, depth, scalar_boundary) in roots {
        let depth = depth + 1;
        // Predicates can turn a failing ANY into empty success at a
        // continuation byte. ASCII-only certificates need no such check.
        let boundary = if scalar_boundary && !failures { " && source.is_char_boundary(offset)" } else { "" };
        let result = if failures { format!("Some(t{id}(source.as_bytes(),offset))") } else { format!("t{id}(source.as_bytes(),offset)") };
        writeln!(
            out,
            "            {r} if remaining_depth>={depth}{boundary}=>{result},"
        )
        .unwrap();
    }
    out.push_str("            _=>None,\n        }\n    }\n");
}
