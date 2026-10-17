use super::{Pair, Rule, WastWalker};
use std::sync::Arc;
use vybe_ast::canon::{ComponentBinary, ComponentBinaryDef};

fn contains_interface(pair: Pair<'_, Rule>) -> bool {
    matches!(pair.as_rule(), Rule::componenttype | Rule::instancetype)
        || pair.into_inner().any(contains_interface)
}

fn install_projection(
    source: Pair<'_, Rule>,
    binary: Arc<ComponentBinary>,
    walker: &mut WastWalker,
) -> Result<(), String> {
    let span = source.as_span();
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
    if children.len() != nested.len() {
        return Err("normalized component tree does not match source declarations".into());
    }
    for (source, binary) in children.into_iter().zip(nested) {
        install_projection(source, binary, walker)?;
    }
    walker
        .component_interface_binaries
        .insert((span.start(), span.end()), binary);
    Ok(())
}

pub(super) fn normalize(source: Pair<'_, Rule>, walker: &mut WastWalker) -> Result<(), String> {
    let span = source.as_span();
    if walker
        .component_interface_binaries
        .contains_key(&(span.start(), span.end()))
        || !contains_interface(source.clone())
    {
        return Ok(());
    }
    // The existing standard text resolver materializes lexical outer aliases.
    // Retain its complete typed binary IR, not just its validation diagnostic.
    let bytes = wat::parse_str(source.as_str())
        .map_err(|error| format!("component interface normalization: {error}"))?;
    let binary = vybe_platform_wasm::read_component_binary(&bytes)?;
    install_projection(source, Arc::new(binary), walker)
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
