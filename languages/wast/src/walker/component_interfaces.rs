use super::{Pair, Rule, WastWalker};
use std::sync::Arc;
use vybe_ast::canon::{ComponentBinary, ComponentBinaryDef};
use super::component_core_scopes::Scope;
use super::component_core_types;

fn contains_interface(pair: Pair<'_, Rule>) -> bool {
    matches!(pair.as_rule(), Rule::componenttype | Rule::instancetype)
        || pair.into_inner().any(contains_interface)
}

fn install_projection(
    source: Pair<'_, Rule>,
    binary: Arc<ComponentBinary>,
    outers: Vec<Scope>,
    walker: &mut WastWalker,
) -> Result<(), String> {
    let span = source.as_span();
    let (projection, child_outers) = component_core_types::index(source.clone(), &binary, outers)?;
    let children: Vec<_> = source
        .into_inner()
        .filter(|child| child.as_rule() == Rule::component_definition)
        .filter_map(|child| child.into_inner().next())
        .filter(|child| child.as_rule() == Rule::component)
        .collect();
    let nested: Vec<_> = binary
        .definitions
        .iter()
        .filter_map(|definition| {
            if let ComponentBinaryDef::NestedComponent(child) = definition {
                Some(Arc::clone(child))
            } else {
                None
            }
        })
        .collect();
    if children.len() != nested.len() || nested.len() != child_outers.len() {
        return Err("normalized component tree does not match source declarations".into());
    }
    for ((source, binary), outers) in children.into_iter().zip(nested).zip(child_outers) {
        install_projection(source, binary, outers, walker)?;
    }
    walker.component_core_rows.extend(projection.rows);
    walker
        .component_interface_binaries
        .insert((span.start(), span.end()), binary);
    Ok(())
}

pub(super) fn normalize(source: Pair<'_, Rule>, walker: &mut WastWalker) -> Result<(), String> {
    let span = source.as_span();
    if super::component_values::contains(source.clone()) {
        if !walker.component_interface_binaries.contains_key(&(span.start(), span.end())) {
            let binary = super::component_values::produce(source.clone())?;
            if super::component_values::concrete(source.clone()) {
                // The binary producer owns its resolved spaces, including
                // anonymous core module types. No executable walk consumes rows.
                walker.component_interface_binaries.insert((span.start(), span.end()), Arc::new(binary));
            } else {
                install_projection(source, Arc::new(binary), Vec::new(), walker)?;
            }
        }
        return Ok(());
    }
    if walker
        .component_interface_binaries
        .contains_key(&(span.start(), span.end()))
        || !(contains_interface(source.clone()) || component_core_types::contains_schema(source.clone()))
    {
        return Ok(());
    }
    // The existing standard text resolver materializes lexical outer aliases.
    // Retain its complete typed binary IR, not just its validation diagnostic.
    let bytes = wat::parse_str(source.as_str())
        .map_err(|error| format!("component interface normalization: {error}"))?;
    let binary = vybe_platform_wasm::read_component_binary(&bytes)?;
    install_projection(source, Arc::new(binary), Vec::new(), walker)
}

pub(super) fn take(source: Pair<'_, Rule>, walker: &mut WastWalker) -> Option<ComponentBinary> {
    let component = if source.as_rule() == Rule::script_cmd {
        source.into_inner().next()?
    } else {
        source
    };
    if component.as_rule() != Rule::component {
        return None;
    }
    let span = component.as_span();
    let binary = walker
        .component_interface_binaries
        .remove(&(span.start(), span.end()))?;
    Some(Arc::try_unwrap(binary).unwrap_or_else(|shared| (*shared).clone()))
}
