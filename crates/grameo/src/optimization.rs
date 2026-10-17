//! Explicit grammar-authored optimization selection. PEG defaults ignore these comments.
use crate::lexer::{Annotation, AnnotationKind};
use crate::{
    CompiledGrammar, Diagnostic, Span,
    islands::{Operator, TrailingTrivia},
    pratt::Fixity,
};

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct OptimizationSelection {
    pub ascii_xid: bool,
    pub nullable_failure_heads: bool,
    pub native_choices: bool,
    pub native_rule_heads: bool,
}
impl CompiledGrammar {
    /// Explicit generated direct-token specialization; absent for ordinary PEG.
    pub fn native_token_compilation(&self) -> bool { self.native_token_compilation }
    /// Explicit opt-in to fully certified negative native-token results.
    pub fn native_token_failures(&self) -> bool { self.native_token_failures }
    /// Explicit ordered-choice targets for generated conservative token-head dispatch.
    /// Default PEG compilation leaves this empty; unknown branch heads retain PEG matching.
    pub fn choice_head_rules(&self)->&[crate::grammar::RuleId]{&self.choice_head_rules}

    pub fn optimization_selection(&self) -> OptimizationSelection {
        self.optimization_selection
    }
}
struct PrattDecl {
    target: String,
    atom: String,
    trailing: TrailingTrivia,
    operator_trivia: Option<String>,
    rhs_atom: Option<String>,
    rhs_operator: Option<String>,
    span: Span,
}
struct OperatorDecl {
    target: String,
    rule: String,
    precedence: u16,
    fixity: Fixity,
    rhs_atom: Option<String>,
    span: Span,
}
fn error(span: Span, message: impl Into<String>) -> Vec<Diagnostic> {
    vec![Diagnostic::new("G015", message, span)]
}
fn property<'a>(word: &'a str, key: &str, span: Span) -> Result<&'a str, Vec<Diagnostic>> {
    word.strip_prefix(key)
        .filter(|v| !v.is_empty())
        .ok_or_else(|| error(span, format!("expected {key}<value>")))
}
pub(crate) fn apply(
    grammar: &mut CompiledGrammar,
    annotations: Vec<Annotation>,
) -> Result<(), Vec<Diagnostic>> {
    let mut selection = OptimizationSelection::default();
    let mut native_tokens = false;
    let mut native_token_failures = false;
    let mut owned_targets:Vec<(String,Span)>=Vec::new();
    let mut empty_suffix_targets:Vec<(String,Span)>=Vec::new();
    let mut choices: Vec<(String,Span)> = Vec::new();
    let mut pratt: Vec<PrattDecl> = Vec::new();
    let mut operators: Vec<OperatorDecl> = Vec::new();
    for a in annotations {
        let words: Vec<_> = a.text.split_ascii_whitespace().collect();
        match a.kind {
            AnnotationKind::Ast => unreachable!("AST annotations have a separate explicit option"),
            AnnotationKind::OwnedCaptureBoundary => {
                if words.is_empty(){return Err(error(a.span,"expected owned_capture_boundary target rule"));}
                for word in words {
                    if owned_targets.iter().any(|(name,_)|name==word){return Err(error(a.span,"duplicate owned_capture_boundary target"));}
                    owned_targets.push((word.into(),a.span));
                }
            }
            AnnotationKind::EmptySuccessHeads=>{
                if words.is_empty(){return Err(error(a.span,"expected empty_success_heads target rules"));}
                for name in words{
                    let rule=grammar.rule_id(name).ok_or_else(||error(a.span,format!("unknown empty-success rule '{name}'")))?;
                    if grammar.empty_success_heads.contains_key(&rule){return Err(error(a.span,"duplicate empty-success target"));}
                    let head=crate::empty_success::certify(grammar,rule).ok_or_else(||error(a.span,"empty-success requires a silent non-Pratt zero-minimum literal-only repetition"))?;
                    grammar.empty_success_heads.insert(rule,head);
                }
            }
            AnnotationKind::EmptySuffixHeads=>{
                if words.is_empty(){return Err(error(a.span,"expected empty_suffix_heads target rules"));}
                for name in words{if empty_suffix_targets.iter().any(|(n,_)|n==name){return Err(error(a.span,"duplicate empty-suffix target"));}empty_suffix_targets.push((name.into(),a.span));}
            }
            AnnotationKind::CodeSpace => {
                let value = match words.as_slice() {
                    ["unit"] => crate::source_units::CodeSpace::Unit,
                    ["request"] => crate::source_units::CodeSpace::Request,
                    _ => return Err(error(a.span, "expected code_space unit|request")),
                };
                if grammar.code_space.replace(value).is_some() {
                    return Err(error(a.span, "duplicate code_space directive"));
                }
            }
            AnnotationKind::SourceUnit => crate::source_units::bind(grammar, &a.text, a.span)?,
            AnnotationKind::ChoiceHeads => {
                if words.is_empty(){return Err(error(a.span,"expected at least one choice_heads target rule"));}
                for word in words{if choices.iter().any(|(name,_)|name==word){return Err(error(a.span,format!("duplicate choice_heads target '{word}'")));}choices.push((word.into(),a.span));}
            }
            AnnotationKind::Optimize => {
                if words.is_empty() {
                    return Err(error(
                        a.span,
                        "expected at least one optimization selection",
                    ));
                }
                for word in words {
                    match word {
                        "native_tokens" => native_tokens = true,
                        "native_token_failures" => { native_tokens = true; native_token_failures = true; },
                        "ascii_xid" => selection.ascii_xid = true,
                        "nullable_failure_heads" => selection.nullable_failure_heads = true,
                        "native_choices" => selection.native_choices = true,
                        "native_rule_heads" => selection.native_rule_heads = true,
                        _ => return Err(error(a.span, format!("unknown optimization '{word}'"))),
                    }
                }
            }
            AnnotationKind::SharedInnerPrefix => {
                if words.len()!=3{return Err(error(a.span,"expected shared_inner_prefix target prefix=<rule> kind=list|index"));}
                let s=crate::inner_prefix::certify(grammar,words[0],property(words[1],"prefix=",a.span)?,property(words[2],"kind=",a.span)?).map_err(|m|error(a.span,m))?;
                if grammar.shared_inner_prefixes.insert(s.choice,s).is_some(){return Err(error(a.span,"duplicate shared-inner-prefix choice"));}
            }
            AnnotationKind::SharedGroupPrefix => {
                if words.len()!=4{return Err(error(a.span,"expected shared_group_prefix target branch=<rule> fallback=<rule> prefix=<rule>"));}
                let s=crate::group_prefix::certify(grammar,words[0],property(words[1],"branch=",a.span)?,property(words[2],"fallback=",a.span)?,property(words[3],"prefix=",a.span)?).map_err(|m|error(a.span,m))?;
                if grammar.shared_group_prefixes.insert(s.choice,s).is_some(){return Err(error(a.span,"duplicate shared-group-prefix choice"));}
            }
            AnnotationKind::SharedWrapperPrefix => {
                if words.len()!=3 { return Err(error(a.span,"expected shared_wrapper_prefix target branch=<rule> fallback=<rule>")); }
                let branch=property(words[1],"branch=",a.span)?;
                let fallback=property(words[2],"fallback=",a.span)?;
                let cert=crate::wrapper_prefix::certify(grammar,words[0],branch,fallback)
                    .map_err(|message|error(a.span,message))?;
                if grammar.shared_wrapper_prefixes.insert(cert.choice,cert).is_some(){return Err(error(a.span,"duplicate shared-wrapper-prefix choice"));}
            }
            AnnotationKind::SharedPrefix => {
                if words.len() != 3 { return Err(error(a.span, "expected shared_prefix target branch=<rule> prefix=<rule>")); }
                let branch = property(words[1], "branch=", a.span)?;
                let prefix = property(words[2], "prefix=", a.span)?;
                let rule_id = |name: &str| grammar.rule_id(name).ok_or_else(|| error(a.span, format!("unknown shared-prefix rule '{name}'")));
                let target = rule_id(words[0])?;
                let branch = rule_id(branch)?;
                let prefix = rule_id(prefix)?;
                use crate::{ExprKind, RuleMode, program::Program};
                fn unwrap(grammar: &CompiledGrammar, mut id: usize) -> usize {
                    while let ExprKind::Group(child) = grammar.syntax().expressions[id].kind { id=child; } id
                }
                let expression = unwrap(grammar, grammar.rule(target).expression);
                if grammar.shared_prefixes.contains_key(&expression) { return Err(error(a.span, "duplicate shared-prefix target")); }
                if grammar.rule(target).mode != RuleMode::Normal || grammar.rule(branch).mode != RuleMode::Normal || grammar.rule(target).scope != grammar.rule(branch).scope {
                    return Err(error(a.span, "shared-prefix target and branch require normal rules in the same scope"));
                }
                let ExprKind::Choice { left, right } = grammar.syntax().expressions[expression].kind else { return Err(error(a.span, "shared-prefix target requires a final ordered choice")); };
                if grammar.reference(unwrap(grammar, right)) != Some(crate::Reference::Rule(prefix)) { return Err(error(a.span, "shared-prefix fallback must call the declared prefix")); }
                let ExprKind::Choice { left: prior, right: wrapper } = grammar.syntax().expressions[unwrap(grammar, left)].kind else { return Err(error(a.span, "shared-prefix branch must follow the prior alternatives")); };
                if grammar.reference(unwrap(grammar, wrapper)) != Some(crate::Reference::Rule(branch)) { return Err(error(a.span, "shared-prefix penultimate alternative must call the declared branch")); }
                let ExprKind::Sequence { left: prefix_expr, right: suffix } = grammar.syntax().expressions[unwrap(grammar, grammar.rule(branch).expression)].kind else { return Err(error(a.span, "shared-prefix branch requires prefix ~ suffix")); };
                if grammar.reference(unwrap(grammar, prefix_expr)) != Some(crate::Reference::Rule(prefix)) { return Err(error(a.span, "shared-prefix branch must begin with the declared prefix")); }
                // Prefix stack effects are retained exactly. The normal inert
                // wrapper adds no stack/environment mutation; restoring the suffix
                // checkpoint retains the deterministic prefix state that the
                // fallback would recreate. This includes balanced fstring delimiters.
                let suffix_rule = match grammar.instruction(unwrap(grammar, suffix)) {
                    crate::program::Instruction::Call(rule) => Some(rule),
                    crate::program::Instruction::Repeat { expression, .. } => match grammar.instruction(unwrap(grammar, expression)) { crate::program::Instruction::Call(rule) => Some(rule), _=>None },
                    _=>None,
                };
                grammar.shared_prefixes.insert(expression, crate::program::SharedPrefixChoice {
                    target, target_expression: grammar.rule(target).expression,
                    branch_expression: grammar.rule(branch).expression, branch_sequence: unwrap(grammar, grammar.rule(branch).expression),
                    left_choice: unwrap(grammar, left), wrapper_call: unwrap(grammar, wrapper), right,
                    scope: grammar.rule(branch).scope, prefix_rule: prefix, suffix_rule,
                    prior, wrapper: branch, prefix: prefix_expr, suffix });
            }
            AnnotationKind::Pratt => {
                if !(3..=6).contains(&words.len()) {
                    return Err(error(
                        a.span,
                        "expected Pratt target atom=<rule> trailing=consume|preserve",
                    ));
                }
                let atom = property(words[1], "atom=", a.span)?;
                let trailing = match property(words[2], "trailing=", a.span)? {
                    "consume" => TrailingTrivia::Consume,
                    "preserve" => TrailingTrivia::Preserve,
                    _ => {
                        return Err(error(
                            a.span,
                            "Pratt trailing trivia must be consume or preserve",
                        ));
                    }
                };
                if let Some(previous) = pratt.iter().find(|p| p.target == words[0]) {
                    let mut errors =
                        error(a.span, format!("duplicate Pratt target '{}'", words[0]));
                    errors[0]
                        .related
                        .push((previous.span, "first Pratt declaration".into()));
                    return Err(errors);
                }
                let mut operator_trivia=None;let mut rhs_atom=None;let mut rhs_operator=None;
                for word in &words[3..] {
                    let (slot,value)=if let Some(value)=word.strip_prefix("operator_trivia=") {(&mut operator_trivia,value)}
                        else if let Some(value)=word.strip_prefix("rhs_operator=") {(&mut rhs_operator,value)}
                        else if let Some(value)=word.strip_prefix("rhs_atom=") {(&mut rhs_atom,value)}
                        else{return Err(error(a.span,format!("unknown Pratt option '{word}'")))};
                    if value.is_empty() || slot.is_some() {return Err(error(a.span,format!("empty or duplicate Pratt option '{word}'")))}
                    *slot=Some(value.to_string());
                }
                pratt.push(PrattDecl {
                    target: words[0].into(),
                    atom: atom.into(),
                    trailing,
                    operator_trivia,
                    rhs_atom,
                    rhs_operator,
                    span: a.span,
                });
            }
            AnnotationKind::Operator => {
                if !(4..=5).contains(&words.len()) {
                    return Err(error(
                        a.span,
                        "expected operator target rule precedence=<u16> fixity=prefix|postfix|infix_left|infix_right",
                    ));
                }
                let precedence = property(words[2], "precedence=", a.span)?
                    .parse()
                    .map_err(|_| error(a.span, "operator precedence must fit u16"))?;
                let fixity = match property(words[3], "fixity=", a.span)? {
                    "prefix" => Fixity::Prefix,
                    "postfix" => Fixity::Postfix,
                    "infix_left" => Fixity::InfixLeft,
                    "infix_right" => Fixity::InfixRight,
                    _ => return Err(error(a.span, "unknown operator fixity")),
                };
                let rhs_atom = words.get(4).map(|word|property(word,"rhs_atom=",a.span).map(str::to_owned)).transpose()?;
                if rhs_atom.is_some() && !matches!(fixity, Fixity::InfixLeft | Fixity::InfixRight) {
                    return Err(error(a.span, "operator RHS atom requires an infix declaration"));
                }
                operators.push(OperatorDecl {
                    target: words[0].into(),
                    rule: words[1].into(),
                    precedence,
                    fixity,
                    rhs_atom,
                    span: a.span,
                });
            }
        }
    }
    for op in &operators {
        if !pratt.iter().any(|p| p.target == op.target) {
            return Err(error(
                op.span,
                format!("operator target '{}' has no Pratt declaration", op.target),
            ));
        }
        if grammar.rule_id(&op.rule).is_none() {
            return Err(error(
                op.span,
                format!("unknown Pratt operator '{}'", op.rule),
            ));
        }
    }
    for p in pratt {
        let bound: Vec<_> = operators
            .iter()
            .filter(|op| op.target == p.target)
            .map(|op| Operator {
                rule: &op.rule,
                precedence: op.precedence,
                fixity: op.fixity,
            })
            .collect();
        grammar
            .bind_pratt(&p.target, &p.atom, &bound, p.trailing)
            .map_err(|mut errors| {
                for e in &mut errors {
                    e.related.push((e.span, "Pratt target rule".into()));
                    e.span = p.span;
                }
                errors
            })?;
        if let Some(rhs)=p.rhs_atom {grammar.set_pratt_rhs_atom(&p.target,&rhs).map_err(|mut errors| {for e in &mut errors {e.related.push((e.span,"Pratt target rule".into()));e.span=p.span;}errors})?;}
        if let Some(operator)=p.rhs_operator {grammar.set_pratt_rhs_operator(&p.target,&operator).map_err(|mut errors| {for e in &mut errors {e.related.push((e.span,"Pratt target rule".into()));e.span=p.span;}errors})?;}
        for op in operators.iter().filter(|op|op.target==p.target) {
            if let Some(rhs)=&op.rhs_atom {grammar.set_pratt_operator_rhs_atom(&p.target,&op.rule,rhs).map_err(|mut errors| {for e in &mut errors {e.related.push((e.span,"Pratt target rule".into()));e.span=op.span;}errors})?;}
        }
        if let Some(trivia)=p.operator_trivia {
            grammar.set_pratt_operator_trivia(&p.target,&trivia).map_err(|mut errors| {for e in &mut errors {e.related.push((e.span,"Pratt target rule".into()));e.span=p.span;}errors})?;
        }
    }
    for (name,span) in choices {
        use crate::program::{Program,Instruction};
        let Some(rule)=grammar.rule_id(&name)else{return Err(error(span,format!("unknown choice_heads target '{name}'")));};
        let spec=grammar.rule(rule);
        if !matches!(spec.mode,crate::RuleMode::Normal|crate::RuleMode::Silent) || grammar.source_pratt(rule).is_some(){return Err(error(span,"choice_heads requires a normal/silent non-Pratt target"));}
        let mut root=spec.expression;
        for _ in 0..grammar.syntax().expressions.len(){if let Instruction::Group(child)=grammar.instruction(root){root=child}else{break;}}
        if !matches!(grammar.instruction(root),Instruction::Choice{..}){return Err(error(span,"choice_heads target must be an ordered choice"));}
        let mut pending=vec![root];let mut count=0;
        while let Some(id)=pending.pop(){if let Instruction::Choice{left,right}=grammar.instruction(id){pending.push(left);pending.push(right);}else{count+=1;}if count+pending.len()>64{return Err(error(span,"choice_heads supports at most 64 original alternatives"));}}
        grammar.choice_head_rules.push(rule);
    }
    grammar.choice_head_rules.sort_unstable();
    for (name,span) in owned_targets {
        use crate::program::Program;
        let rule=grammar.rule_id(&name).ok_or_else(||error(span,format!("unknown owned capture rule '{name}'")))?;
        let definition=grammar.rule(rule);
        let scope=grammar.scope(definition.scope);
        if definition.mode!=crate::RuleMode::Normal || scope.whitespace==Some(rule) || scope.comment==Some(rule) || grammar.source_pratt(rule).is_some() {
            return Err(error(span,"owned capture target requires a normal non-trivia non-Pratt rule"));
        }
        let rules=owned_capture_closure(grammar,rule);
        grammar.owned_capture_boundaries.insert(rule,rules);
    }
    for (name,span)in empty_suffix_targets{
        let rule=grammar.rule_id(&name).ok_or_else(||error(span,format!("unknown empty-suffix rule {name}")))?;
        let h=crate::empty_suffix::certify(grammar,rule).ok_or_else(||error(span,"empty-suffix requires unique optional tail with pure nullable prefix, mandatory literal head and known trivia heads"))?;
        if grammar.empty_suffix_heads.insert(h.repetition,h).is_some(){return Err(error(span,"duplicate empty-suffix target"));}
    }
    for &rule in grammar.empty_success_heads.keys(){
        if crate::empty_success::certify(grammar,rule).is_none(){return Err(error(grammar.syntax().rules[rule].span,"empty-success target cannot be reconfigured as Pratt"));}
    }
    grammar.native_token_compilation = native_tokens;
    grammar.native_token_failures = native_token_failures;
    grammar.optimization_selection = selection;
    Ok(())
}

/// Complete conservative call closure, including implicit scoped trivia and
/// configured Pratt atom/operator/RHS entries. Original expressions remain intact.
fn owned_capture_closure(grammar:&CompiledGrammar,root:usize)->Vec<usize>{
    use crate::program::{Program,Instruction};
    let mut rules=std::collections::BTreeSet::new();
    let mut expressions=std::collections::HashSet::new();
    let mut pending_rules=vec![root];let mut pending_expr=Vec::new();
    loop {
        while let Some(rule)=pending_rules.pop(){
            if !rules.insert(rule){continue;}
            let definition=grammar.rule(rule);pending_expr.push(definition.expression);
            let scope=grammar.scope(definition.scope);
            pending_rules.extend(scope.whitespace);pending_rules.extend(scope.comment);
            if let Some(p)=grammar.pratt[rule].as_ref(){
                pending_rules.push(p.atom);pending_rules.extend(p.operators.iter().map(|o|o.rule));
                pending_rules.extend(p.rhs_atom);pending_rules.extend(p.rhs_operator);pending_rules.extend(p.operator_trivia);
                pending_rules.extend(p.operator_rhs_atoms.iter().map(|(_,r)|*r));
            }
        }
        let Some(expr)=pending_expr.pop() else {break;};
        if !expressions.insert(expr){continue;}
        match grammar.instruction(expr){
            Instruction::Call(rule)=>pending_rules.push(rule),
            Instruction::Sequence{left,right}|Instruction::Choice{left,right}=>{pending_expr.push(left);pending_expr.push(right);},
            Instruction::Repeat{expression,..}|Instruction::Predicate{expression,..}|Instruction::Group(expression)|Instruction::Push(expression)=>pending_expr.push(expression),
            _=>{}
        }
    }
    rules.into_iter().collect()
}
