//! Bind the existing parser's source names to validated, typed normalized schemas.
use super::{CoreTypeKind, Pair, Rule, WastWalker, find_rule};
use super::component_core_scopes::{Scope, Slot};
use std::collections::HashMap;
use vybe_ast::canon::{ComponentBinary, ComponentBinaryDef as Def,
    ComponentCoreCompositeType as Composite, ComponentCoreType, ComponentCoreValueType as Value,
    ComponentCoreValType as Scalar};

type Span = (usize, usize);

#[derive(Clone, Debug)]
pub(super) struct Row {
    pub definition: Def,
    pub slots: Vec<Slot>,
    pub names: Vec<Option<String>>,
}

pub(super) struct Projection {
    pub rows: HashMap<Span, Row>,
}

fn span(pair: &Pair<'_, Rule>) -> Span {
    let span = pair.as_span();
    (span.start(), span.end())
}

fn binder(pair: &Pair<'_, Rule>) -> Option<String> {
    pair.clone().into_inner().find(|part| part.as_rule() == Rule::id)
        .map(|part| part.as_str()[1..].to_owned())
}

pub(super) fn is_outer_alias(pair: &Pair<'_, Rule>) -> bool {
    pair.as_rule() == Rule::comp_alias && find_rule(pair, Rule::core_outer_type_sort).is_some()
}

pub(super) fn contains_schema(pair: Pair<'_, Rule>) -> bool {
    matches!(pair.as_rule(), Rule::core_type_def | Rule::core_rec_group)
        || is_outer_alias(&pair) || pair.into_inner().any(contains_schema)
}

/// Return the child component's lexical core scopes at its declaration point.
pub(super) fn index(source: Pair<'_, Rule>, binary: &ComponentBinary, outers: Vec<Scope>)
    -> Result<(Projection, Vec<Vec<Scope>>), String> {
    let mut scope = Scope::default();
    let mut normalized = Vec::new();
    let mut children = Vec::new();
    for definition in &binary.definitions {
        if let Some(slots) = scope.declaration(definition, &outers)? {
            normalized.push((definition.clone(), slots));
        }
        if matches!(definition, Def::NestedComponent(_)) {
            let mut lexical = outers.clone();
            lexical.push(scope.clone());
            children.push(lexical);
        }
    }
    let mut normalized = normalized.into_iter();
    let mut rows = HashMap::new();
    for wrapped in source.into_inner().filter(|part| part.as_rule() == Rule::component_definition) {
        let declaration = wrapped.into_inner().next().ok_or("empty component source declaration")?;
        if !matches!(declaration.as_rule(), Rule::core_type_def | Rule::core_rec_group)
            && !is_outer_alias(&declaration) { continue; }
        let (definition, slots) = normalized.next().ok_or("normalized core type declaration is missing")?;
        let names = if declaration.as_rule() == Rule::core_rec_group {
            declaration.clone().into_inner().filter(|part| part.as_rule() == Rule::core_member_type)
                .map(|part| binder(&part)).collect::<Vec<_>>()
        } else { vec![binder(&declaration)] };
        if names.len() != slots.len() {
            return Err("source recursive group does not match normalized core type slots".into());
        }
        let expected_alias = is_outer_alias(&declaration);
        if expected_alias != definition.normalized_outer_alias().is_some_and(|alias|
            alias.kind == vybe_ast::canon::ComponentOuterAliasKind::CoreType) {
            return Err("source core declaration sort does not match normalized metadata".into());
        }
        if declaration.as_rule() == Rule::core_rec_group
            && !matches!(&definition, Def::CoreType(ComponentCoreType::Rec(group)) if group.explicit) {
            return Err("source recursive boundary is missing from normalized metadata".into());
        }
        rows.insert(span(&declaration), Row { definition, slots, names });
    }
    if normalized.next().is_some() {
        return Err("normalized core type declarations have no source binders".into());
    }
    Ok((Projection { rows }, children))
}

pub(super) fn row<'a>(walker: &'a WastWalker, source: &Pair<'_, Rule>) -> Result<&'a Row, String> {
    walker.component_core_rows.get(&span(source))
        .ok_or_else(|| "core type has no retained typed source declaration".into())
}

fn kind(slot: &Slot) -> CoreTypeKind {
    if matches!(slot.definition.as_ref(), ComponentCoreType::Module(_)) { return CoreTypeKind::Module; }
    if let Some(ty) = slot.subtype() {
        if let Composite::Func { params, results } = &ty.composite {
            if results.is_empty() {
                return match params.as_slice() {
                    [Value::Scalar(value @ (Scalar::I32 | Scalar::I64))] => CoreTypeKind::Worker(*value),
                    _ => CoreTypeKind::Other,
                };
            }
        }
    }
    CoreTypeKind::Other
}

pub(super) fn walk(walker: &mut WastWalker, source: Pair<'_, Rule>) -> Result<(), String> {
    let row = row(walker, &source)?.clone();
    for (name, slot) in row.names.iter().zip(&row.slots) {
        let index = walker.core_type_space;
        walker.core_type_space = index.checked_add(1).ok_or("core type index overflow")?;
        walker.core_type_kinds.push(kind(slot));
        if let Some(name) = name {
            if walker.core_type_index.insert(name.clone(), index).is_some() {
                return Err(format!("duplicate core type name ${name}"));
            }
        }
    }
    Ok(())
}

pub(super) fn project_binary(row: &Row, names: &mut HashMap<String, u32>, count: &mut u32)
    -> Result<Def, String> {
    for name in &row.names {
        let index = *count;
        *count = count.checked_add(1).ok_or("core type index overflow")?;
        if let Some(name) = name {
            if names.insert(name.clone(), index).is_some() { return Err(format!("duplicate core type name ${name}")); }
        }
    }
    Ok(row.definition.clone())
}

#[cfg(test)]
#[path = "component_core_type_tests.rs"]
mod tests;
