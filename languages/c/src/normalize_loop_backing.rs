//! Version typed scalar loops over numeric backing, retaining managed fallback.

use std::collections::{HashMap, HashSet};
use vybe_ast::{
    BinOp, BindingPattern, ExprKind, Expression, Literal, Param, PassBy, PlaceExpr, Statement,
    StmtKind, StorageAccess, UnaryOp,
};

fn ident(value: &Expression) -> Option<&str> {
    if let ExprKind::Ident(name) = &value.kind {
        Some(name)
    } else {
        None
    }
}

fn declarations(
    body: &mut Vec<Statement>,
    facts: &mut HashMap<String, bool>,
    duplicates: &mut HashSet<String>,
) {
    for statement in body {
        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                if let BindingPattern::Ident(name) = &declaration.pattern {
                    let ordinary = declaration
                        .type_hint
                        .as_ref()
                        .is_some_and(|hint| hint.access == Some(StorageAccess::Ordinary));
                    if facts.insert(name.clone(), ordinary).is_some() {
                        duplicates.insert(name.clone());
                    }
                }
            }
        }
        crate::normalize_affine_bytes::children(statement, &mut |body| {
            declarations(body, facts, duplicates)
        });
    }
}

fn addressed(body: &[Statement]) -> HashSet<String> {
    let mut result = HashSet::new();
    for statement in body {
        statement.walk_exprs(&mut |value| match &value.kind {
            ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr,
            } => {
                if let Some(name) = ident(expr) {
                    result.insert(name.into());
                }
            }
            ExprKind::RefOf(place) => {
                if let PlaceExpr::Ident(name) = &**place {
                    result.insert(name.clone());
                }
            }
            _ => {}
        });
    }
    result
}

fn numeric(value: &Expression, known: &HashSet<String>) -> bool {
    match &value.kind {
        ExprKind::Ident(name) => known.contains(name),
        ExprKind::Lit(Literal::Int(_) | Literal::Float(_)) => true,
        ExprKind::Binary {
            op: BinOp::Integer(operation, _),
            ..
        } if !operation.is_comparison() => true,
        ExprKind::Binary {
            op:
                BinOp::Add
                | BinOp::Sub
                | BinOp::Mul
                | BinOp::BitAnd
                | BinOp::BitOr
                | BinOp::BitXor
                | BinOp::Shl
                | BinOp::Shr,
            left,
            right,
        } => numeric(left, known) && numeric(right, known),
        ExprKind::Cast { type_name, .. } => matches!(
            type_name.as_str(),
            "int"
                | "integer"
                | "int32"
                | "uint32"
                | "uint8"
                | "uint16"
                | "int8"
                | "int16"
                | "double"
        ),
        ExprKind::Call { callee, .. } => matches!(
            ident(callee),
            Some(
                "__c_ptr_i32_load8_u"
                    | "__c_ptr_i32_load8_s"
                    | "__c_ptr_i32_load16_u"
                    | "__c_ptr_i32_load16_s"
                    | "__c_ptr_i32_load"
                    | "__c_unsigned32_wrap"
                    | "__c_unsigned32_from_bits"
            )
        ),
        ExprKind::Sequence(values) => values.last().is_some_and(|value| numeric(value, known)),
        ExprKind::Assign { value, .. } => numeric(value, known),
        ExprKind::Unary {
            op:
                UnaryOp::PostInc | UnaryOp::PreInc | UnaryOp::PostDec | UnaryOp::PreDec | UnaryOp::Neg,
            expr,
        } => numeric(expr, known),
        _ => false,
    }
}

fn boolean(value: &Expression) -> Option<bool> {
    if let ExprKind::Lit(Literal::Bool(value)) = value.kind {
        Some(value)
    } else {
        None
    }
}

// A folded branch is owned by a parent that is about to be discarded. Move
// its kind instead of cloning the complete subtree; the parent's span remains
// unchanged, matching the existing folding behavior.
fn take_folded_kind(value: &mut Expression) -> ExprKind {
    std::mem::replace(&mut value.kind, ExprKind::Lit(Literal::Int(0)))
}

fn specialize(value: &mut Expression, known: &mut HashSet<String>) {
    match &mut value.kind {
        ExprKind::Ternary { cond, then, else_ } => {
            specialize(cond, known);
            if let Some(condition) = boolean(cond) {
                let selected = if condition {
                    then.as_mut()
                } else {
                    else_.as_mut()
                };
                specialize(selected, known);
                value.kind = take_folded_kind(selected);
            } else {
                let mut lhs = known.clone();
                let mut rhs = known.clone();
                specialize(then, &mut lhs);
                specialize(else_, &mut rhs);
                *known = lhs.intersection(&rhs).cloned().collect();
            }
        }
        ExprKind::Binary { op, left, right } => {
            specialize(left, known);
            if matches!(op, BinOp::And | BinOp::Or) {
                if let Some(lhs) = boolean(left) {
                    if (*op == BinOp::And && !lhs) || (*op == BinOp::Or && lhs) {
                        value.kind = ExprKind::Lit(Literal::Bool(lhs));
                        return;
                    }
                    specialize(right, known);
                    value.kind = take_folded_kind(right);
                    return;
                }
                let mut rhs = known.clone();
                specialize(right, &mut rhs);
                *known = known.intersection(&rhs).cloned().collect();
            } else {
                specialize(right, known);
            }
            if matches!(
                op,
                BinOp::Eq | BinOp::StrictEq | BinOp::NotEq | BinOp::StrictNotEq
            ) {
                if let (ExprKind::Lit(Literal::Str(lhs)), ExprKind::Lit(Literal::Str(rhs))) =
                    (&left.kind, &right.kind)
                {
                    let equal = lhs == rhs;
                    value.kind = ExprKind::Lit(Literal::Bool(
                        if matches!(op, BinOp::NotEq | BinOp::StrictNotEq) {
                            !equal
                        } else {
                            equal
                        },
                    ));
                }
            }
        }
        ExprKind::Sequence(values) => {
            for value in values {
                specialize(value, known);
            }
        }
        ExprKind::Assign { target, value } => {
            specialize(value, known);
            if let Some(name) = ident(target) {
                if numeric(value, known) {
                    known.insert(name.into());
                } else {
                    known.remove(name);
                }
            }
        }
        ExprKind::IsType { expr, type_name } => {
            specialize(expr, known);
            if ident(expr).is_some_and(|name| known.contains(name))
                && matches!(type_name.as_str(), "number" | "object" | "string")
            {
                value.kind = ExprKind::Lit(Literal::Bool(type_name == "number"));
            }
        }
        ExprKind::Unary {
            op: UnaryOp::Typeof,
            expr,
        } => {
            specialize(expr, known);
            if ident(expr).is_some_and(|name| known.contains(name)) {
                value.kind = ExprKind::Lit(Literal::Str("number".into()));
            }
        }
        ExprKind::Cast { expr, type_name } => {
            specialize(expr, known);
            if (type_name.contains('*') || type_name == "pointerCast") && numeric(expr, known) {
                value.kind = expr.kind.clone();
            }
        }
        ExprKind::Unary { expr, .. } => specialize(expr, known),
        ExprKind::Call { args, .. } => {
            for arg in args {
                specialize(&mut arg.value, known);
            }
        }
        _ => {}
    }
}

fn safe(value: &Expression) -> bool {
    let mut valid = true;
    value.walk_exprs(&mut |value| match &value.kind {
        ExprKind::Ident(_)
        | ExprKind::Lit(_)
        | ExprKind::Binary { .. }
        | ExprKind::Sequence(_)
        | ExprKind::Ternary { .. }
        | ExprKind::Cast { .. }
        | ExprKind::IsType { .. } => {}
        ExprKind::Assign { target, .. } if ident(target).is_some() => {}
        ExprKind::Unary {
            op:
                UnaryOp::PostInc
                | UnaryOp::PreInc
                | UnaryOp::PostDec
                | UnaryOp::PreDec
                | UnaryOp::Neg
                | UnaryOp::Not
                | UnaryOp::Typeof,
            ..
        } => {}
        ExprKind::Call {
            callee,
            args,
            optional: false,
        } if matches!(
            ident(callee),
            Some(
                "__c_ptr_i32_load8_u"
                    | "__c_ptr_i32_load8_s"
                    | "__c_ptr_i32_load16_u"
                    | "__c_ptr_i32_load16_s"
                    | "__c_ptr_i32_load"
                    | "__c_ptr_i32_store8"
                    | "__c_ptr_byte_store"
                    | "__c_unsigned32_wrap"
                    | "__c_unsigned32_from_bits"
            )
        ) && args
            .iter()
            .all(|arg| !arg.by_ref && !arg.spread && arg.name.is_none()) => {}
        _ => valid = false,
    });
    valid
}

fn version(statement: &mut Statement, facts: &HashMap<String, bool>, forbidden: &HashSet<String>) {
    let (cond, body) = match &statement.kind {
        StmtKind::While {
            cond,
            body,
            else_body: None,
        }
        | StmtKind::DoWhile {
            cond,
            body,
            until: false,
        } => (cond, body),
        _ => return,
    };
    let mut loop_bindings = HashSet::new();
    for statement in body {
        match &statement.kind {
            StmtKind::Expr(_) => {}
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        return;
                    };
                    if declaration.with_events
                        || declaration.array_bounds.is_some()
                        || facts.get(name) != Some(&true)
                        || forbidden.contains(name)
                    {
                        return;
                    }
                    loop_bindings.insert(name.clone());
                }
            }
            _ => return,
        }
    }
    let mut candidates = HashSet::new();
    let mut unproved_access = false;
    statement.walk_exprs(&mut |value| {
        if let Some(name) = ident(value) {
            unproved_access |= facts.get(name) == Some(&false) || forbidden.contains(name);
        }
        if let ExprKind::IsType { expr, type_name } = &value.kind {
            if type_name == "number" {
                if let Some(name) = ident(expr).filter(|name| {
                    facts.get(*name) == Some(&true)
                        && !forbidden.contains(*name)
                        && !loop_bindings.contains(*name)
                }) {
                    candidates.insert(name.to_string());
                }
            }
        }
    });
    if candidates.is_empty() || unproved_access {
        return;
    }
    let original = statement.clone();
    let mut fast = original.clone();
    let mut known = candidates.clone();
    let pretest = matches!(fast.kind, StmtKind::While { .. });
    let (fast_cond, fast_body) = match &mut fast.kind {
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => (cond, body),
        _ => unreachable!(),
    };
    if pretest {
        specialize(fast_cond, &mut known);
        if !safe(fast_cond) {
            return;
        }
    }
    for statement in fast_body {
        match &mut statement.kind {
            StmtKind::Expr(value) => {
                specialize(value, &mut known);
                if !safe(value) {
                    return;
                }
            }
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        unreachable!()
                    };
                    known.remove(name);
                    if let Some(value) = &mut declaration.init {
                        specialize(value, &mut known);
                        if !safe(value) {
                            return;
                        }
                        if numeric(value, &known) {
                            known.insert(name.clone());
                        }
                    }
                }
            }
            _ => unreachable!(),
        }
    }
    if !pretest {
        specialize(fast_cond, &mut known);
    }
    if !safe(fast_cond) || !candidates.is_subset(&known) {
        return;
    }
    // No new volatile/addressed binding reads may be introduced at loop entry.
    let mut names = candidates.into_iter().collect::<Vec<_>>();
    names.sort();
    let guard = names
        .into_iter()
        .map(|name| {
            Expression::new(ExprKind::IsType {
                expr: Box::new(Expression::ident(&name)),
                type_name: "number".into(),
            })
        })
        .reduce(|left, right| {
            Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(left),
                right: Box::new(right),
            })
        })
        .unwrap();
    let _ = (cond, body);
    statement.kind = StmtKind::If {
        cond: guard,
        then_body: vec![fast],
        elifs: vec![],
        else_body: Some(vec![original]),
    };
}

fn loops(body: &mut Vec<Statement>, facts: &HashMap<String, bool>, forbidden: &HashSet<String>) {
    for statement in body {
        if let StmtKind::If {
            then_body,
            else_body: Some(other),
            ..
        } = &statement.kind
        {
            if then_body.len() == 1
                && other.len() == 1
                && matches!(
                    then_body[0].kind,
                    StmtKind::While { .. } | StmtKind::DoWhile { .. }
                )
                && matches!(
                    other[0].kind,
                    StmtKind::While { .. } | StmtKind::DoWhile { .. }
                )
            {
                continue;
            }
        }
        crate::normalize_affine_bytes::children(statement, &mut |body| {
            loops(body, facts, forbidden)
        });
        version(statement, facts, forbidden);
    }
}

pub fn normalize(body: &mut Vec<Statement>) {
    let mut globals = HashMap::new();
    for statement in body.iter_mut() {
        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                if let BindingPattern::Ident(name) = &declaration.pattern {
                    let ordinary = declaration
                        .type_hint
                        .as_ref()
                        .is_some_and(|hint| hint.access == Some(StorageAccess::Ordinary));
                    globals
                        .entry(name.clone())
                        .and_modify(|previous| *previous &= ordinary)
                        .or_insert(ordinary);
                }
            }
        }
    }
    let exposed = addressed(body)
        .into_iter()
        .filter(|name| globals.contains_key(name))
        .collect::<HashSet<_>>();
    for statement in body {
        if let StmtKind::FunctionDecl { params, body, .. } = &mut statement.kind {
            let mut facts = HashMap::new();
            let mut forbidden = exposed.clone();
            for Param {
                name,
                type_hint,
                pass_by,
                ..
            } in params
            {
                facts.insert(
                    name.clone(),
                    type_hint
                        .as_ref()
                        .is_some_and(|hint| hint.access == Some(StorageAccess::Ordinary)),
                );
                if *pass_by != PassBy::Value {
                    forbidden.insert(name.clone());
                }
            }
            declarations(body, &mut facts, &mut forbidden);
            for (name, ordinary) in &globals {
                facts.entry(name.clone()).or_insert(*ordinary);
            }
            forbidden.extend(addressed(body));
            let mut captures = false;
            for statement in body.iter() {
                statement.walk_exprs(&mut |value| {
                    captures |= matches!(&value.kind, ExprKind::Lambda { captures, .. } if !captures.is_empty());
                });
            }
            if !captures {
                loops(body, &facts, &forbidden);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn versions(source: &str) -> (usize, vybe_ast::Module) {
        let module = crate::parse(source).unwrap();
        let count = module
            .body
            .iter()
            .filter_map(|statement| {
                if let StmtKind::FunctionDecl { body, .. } = &statement.kind {
                    Some(body)
                } else {
                    None
                }
            })
            .flatten()
            .filter(|statement| {
                matches!(&statement.kind, StmtKind::If {
            then_body, else_body: Some(other), ..
        } if then_body.len() == 1 && other.len() == 1
            && matches!(then_body[0].kind, StmtKind::DoWhile { .. })
            && matches!(other[0].kind, StmtKind::DoWhile { .. }))
            })
            .count();
        (count, module)
    }

    #[test]
    fn versions_typed_scalar_lookup_loop_and_restores_pointer_binding_hints() {
        let (count, mut module) = versions(
            "typedef unsigned char byte;
            byte *a; byte *b;
            void raster(byte *d, int n, unsigned int p, unsigned int step) {
                do { *d++ = b[a[(p >> 26)]]; p += step; } while(n--);
            }",
        );
        assert_eq!(count, 1);
        let params = module.body.iter().find_map(|statement| match &statement.kind {
            StmtKind::FunctionDecl { name, params, .. } if name == "raster" => Some(params),
            _ => None,
        }).unwrap();
        assert!(params.iter().find(|param| param.name == "d").unwrap().type_hint.is_none());
        let before = format!("{:?}", module.body);
        normalize(&mut module.body);
        assert_eq!(before, format!("{:?}", module.body));
    }

    #[test]
    fn rejects_volatile_typedef_pointers_exposed_cursors_and_unknown_callbacks() {
        for source in [
            "void f(volatile unsigned char *d, unsigned char *s, int n) {
                do { *d++ = *s; } while(n--); }",
            "typedef volatile unsigned char byte;
             void f(byte *d, unsigned char *s, int n) {
                do { *d++ = *s; } while(n--); }",
            "void f(unsigned char *d, unsigned char *s, int n) {
                unsigned char **alias = &d; do { *d++ = *s; } while(n--); }",
            "extern void callback(void);
             void f(unsigned char *d, unsigned char *s, int n) {
                do { *d++ = *s; callback(); } while(n--); }",
        ] {
            assert_eq!(versions(source).0, 0, "{source}");
        }
    }

    #[test]
    fn versions_ordinary_loop_local_index_without_reading_it_at_entry() {
        let (count, module) = versions(
            "typedef unsigned char byte; byte *a; byte *b;
             void raster(byte *d, int n, unsigned int p, unsigned int step) {
                 do { unsigned int index = ((p >> 4) & 4032) | (p >> 26);
                      *d++ = b[a[index]]; p += step; } while(n--);
             }",
        );
        assert_eq!(count, 1);
        for statement in &module.body {
            if let StmtKind::FunctionDecl { body, .. } = &statement.kind {
                for statement in body {
                    if let StmtKind::If {
                        cond, then_body, ..
                    } = &statement.kind
                    {
                        if matches!(
                            then_body.first().map(|statement| &statement.kind),
                            Some(StmtKind::DoWhile { .. })
                        ) {
                            cond.walk_exprs(&mut |value| {
                                assert!(!ident(value).is_some_and(|name| name.ends_with("_index")));
                            });
                        }
                    }
                }
            }
        }
    }
}
