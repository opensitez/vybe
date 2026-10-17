//! Explicit source-language Pratt bindings over existing grammar rule names.
//! These are optimization/semantic contracts, never inferred from PEG order.
use crate::{
    CompiledGrammar, Diagnostic, Span,
    pratt::Fixity,
    program::{BoundOperator, OwnedPratt, Program},
};

#[derive(Debug, Clone, Copy)]
pub struct Operator<'a> {
    pub rule: &'a str,
    pub precedence: u16,
    pub fixity: Fixity,
}
/// Match the original expression rule's trailing trivia behavior explicitly.
/// A normal `atom ~ operator_tail*` consumes trivia before its optional tail;
/// a bare atom or an atomic expression generally preserves it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TrailingTrivia {
    Preserve,
    Consume,
}
impl CompiledGrammar {
    /// Opt into Pratt for recognition and cooperating direct-AST builders.
    /// The grammar author certifies the table describes the target rule's
    /// expression language. Capture mode stays on the original PEG grammar.
    /// Operator rules must be lexical: their semantic hooks are suppressed.
    /// Atoms use the ordinary rule engine and may read or modify the grammar
    /// stack. Its transaction journal restores these effects on backtracking.
    pub fn bind_pratt(
        &mut self,
        target: &str,
        atom: &str,
        operators: &[Operator<'_>],
        trailing_trivia: TrailingTrivia,
    ) -> Result<(), Vec<Diagnostic>> {
        let span = self
            .rule_id(target)
            .map_or(Span { start: 0, end: 0 }, |id| {
                self.syntax().rules[id].name_span
            });
        let error = |message: String| vec![Diagnostic::new("G014", message, span)];
        let target = self
            .rule_id(target)
            .ok_or_else(|| error(format!("unknown Pratt target '{target}'")))?;
        let atom = self
            .rule_id(atom)
            .ok_or_else(|| error(format!("unknown Pratt atom '{atom}'")))?;
        if target == atom || operators.len() > 128 {
            return Err(error(
                "Pratt atom must differ from target; at most 128 operator bindings are allowed"
                    .into(),
            ));
        }
        let mut bound = Vec::new();
        let mut positions = std::collections::HashSet::new();
        for op in operators {
            let rule = self
                .rule_id(op.rule)
                .ok_or_else(|| error(format!("unknown Pratt operator '{}'", op.rule)))?;
            let prefix = op.fixity == Fixity::Prefix;
            if !positions.insert((rule, prefix)) {
                return Err(error(format!("ambiguous Pratt operator '{}'", op.rule)));
            }
            bound.push(BoundOperator {
                rule,
                precedence: op.precedence,
                fixity: op.fixity,
            });
        }
        for rule in std::iter::once(target)
            .chain(std::iter::once(atom))
            .chain(bound.iter().map(|op| op.rule))
        {
            let facts = self.analysis().rules[rule];
            if facts.may_succeed_empty {
                return Err(error(format!(
                    "Pratt rule '{}' is nullable/uncertain; use the ordinary backend",
                    self.rule(rule).name
                )));
            }
        }
        // Only atoms execute with the ordinary semantic/stack environment.
        // Operators execute with suppressed hooks and must remain lexical;
        // their stack effects must not leak across speculative reductions.
        for operator in &bound {
            if self.analysis().rules[operator.rule].effects.is_stack_dependent() {
                return Err(error(format!(
                    "Pratt operator '{}' is stack-dependent; use a lexical operator rule",
                    self.rule(operator.rule).name
                )));
            }
        }
        self.empty_suffix_heads.clear();
        self.owned_capture_boundaries.clear(); // Configured reachable rules invalidate capture-observer closure.
        self.pratt[target] = Some(OwnedPratt {
            operator_rhs_atoms: Vec::new(),
            rhs_atom: None,
            rhs_operator: None,
            operator_trivia: None,
            atom,
            operators: bound,
            trailing_trivia,
        });
        Ok(())
    }
    /// Explicitly bind the source grammar separator after an infix operator.
    /// It executes as a normal rule, retaining its captures, scopes, hooks and
    /// grammar-stack transactions. Nullable separators are permitted.
    pub fn set_pratt_operator_trivia(&mut self,target:&str,trivia:&str)->Result<(),Vec<Diagnostic>> {
        let span=self.rule_id(target).map_or(Span{start:0,end:0},|id|self.syntax().rules[id].name_span);
        let error=|text:String|vec![Diagnostic::new("G014",text,span)];
        let target_id=self.rule_id(target).ok_or_else(||error(format!("unknown Pratt target '{target}'")))?;
        let trivia_id=self.rule_id(trivia).ok_or_else(||error(format!("unknown Pratt operator trivia '{trivia}'")))?;
        let spec=self.pratt[target_id].as_mut().ok_or_else(||error(format!("Pratt target '{target}' has no binding")))?;
        if trivia_id==target_id || trivia_id==spec.atom {return Err(error("Pratt operator trivia must differ from the target and atom".into()));}
        self.empty_suffix_heads.clear();
        self.owned_capture_boundaries.clear();
        spec.operator_trivia=Some(trivia_id);
        Ok(())
    }

    pub fn set_pratt_rhs_atom(&mut self,target:&str,rhs:&str)->Result<(),Vec<Diagnostic>> {
        let span=self.rule_id(target).map_or(Span{start:0,end:0},|id|self.syntax().rules[id].name_span);
        let error=|message:String|vec![Diagnostic::new("G014",message,span)];
        let target_id=self.rule_id(target).ok_or_else(||error(format!("unknown Pratt target '{target}'")))?;
        let rhs_id=self.rule_id(rhs).ok_or_else(||error(format!("unknown Pratt RHS atom '{rhs}'")))?;
        if rhs_id==target_id || self.analysis().rules[rhs_id].may_succeed_empty {return Err(error("Pratt RHS atom must differ from target and be non-nullable".into()));}
        let spec=self.pratt[target_id].as_mut().ok_or_else(||error(format!("Pratt target '{target}' has no binding")))?;
        self.empty_suffix_heads.clear();
        self.owned_capture_boundaries.clear();
        spec.rhs_atom=Some(rhs_id);Ok(())
    }

    pub fn set_pratt_rhs_operator(&mut self,target:&str,operator:&str)->Result<(),Vec<Diagnostic>> {
        let span=self.rule_id(target).map_or(Span{start:0,end:0},|id|self.syntax().rules[id].name_span);
        let error=|message:String|vec![Diagnostic::new("G014",message,span)];
        let target_id=self.rule_id(target).ok_or_else(||error(format!("unknown Pratt target '{target}'")))?;
        let op_id=self.rule_id(operator).ok_or_else(||error(format!("unknown Pratt RHS operator '{operator}'")))?;
        let spec=self.pratt[target_id].as_mut().ok_or_else(||error(format!("Pratt target '{target}' has no binding")))?;
        if spec.rhs_atom.is_none() || !spec.operators.iter().any(|op|op.rule==op_id && matches!(op.fixity,Fixity::InfixLeft|Fixity::InfixRight)) {return Err(error("RHS operator requires RHS atom and a bound infix operator".into()));}
        self.empty_suffix_heads.clear();
        self.owned_capture_boundaries.clear();
        spec.rhs_operator=Some(op_id);Ok(())
    }


    /// Select an operand rule for one registered infix operator.
    pub fn set_pratt_operator_rhs_atom(&mut self,target:&str,operator:&str,atom:&str)->Result<(),Vec<Diagnostic>> {
        let span=self.rule_id(target).map_or(Span{start:0,end:0},|id|self.syntax().rules[id].name_span);
        let error=|message:String|vec![Diagnostic::new("G014",message,span)];
        let target_id=self.rule_id(target).ok_or_else(||error(format!("unknown Pratt target '{target}'")))?;
        let op_id=self.rule_id(operator).ok_or_else(||error(format!("unknown Pratt operator '{operator}'")))?;
        let atom_id=self.rule_id(atom).ok_or_else(||error(format!("unknown Pratt operand '{atom}'")))?;
        if atom_id==target_id || self.analysis().rules[atom_id].may_succeed_empty {return Err(error("Pratt operator operand must differ from target and be non-nullable".into()));}
        let spec=self.pratt[target_id].as_mut().ok_or_else(||error(format!("Pratt target '{target}' has no binding")))?;
        if !spec.operators.iter().any(|op|op.rule==op_id && matches!(op.fixity,Fixity::InfixLeft|Fixity::InfixRight)) {return Err(error("Pratt operator operand requires a bound infix operator".into()));}
        if spec.operator_rhs_atoms.iter().any(|&(op,_)|op==op_id) {return Err(error("duplicate Pratt operator operand".into()));}
        self.empty_suffix_heads.clear();
        self.owned_capture_boundaries.clear();
        spec.operator_rhs_atoms.push((op_id,atom_id));Ok(())
    }
}
