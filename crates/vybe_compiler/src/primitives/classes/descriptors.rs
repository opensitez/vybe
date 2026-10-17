//! Structural descriptor inheritance, shared by all source class models.

use vybe_runtime::chunk::TypeEntry;

/// Link descriptor supertypes only after every descriptor row exists. Class
/// declaration order and diagnostic names do not determine type identity.
pub(super) fn link_supertypes(types: &mut [TypeEntry]) -> Result<(), String> {
    let mut descriptor_for_class = vec![None; types.len()];
    for (index, row) in types.iter().enumerate() {
        if let Some(class_index) = row.describes_index.checked_sub(1) {
            let slot = descriptor_for_class.get_mut(class_index as usize)
                .ok_or_else(|| "descriptor refers to an undeclared class type".to_owned())?;
            *slot = Some(u16::try_from(index + 1)
                .map_err(|_| "descriptor type index exceeds the class type space".to_owned())?);
        }
    }

    // Compute first, then publish: malformed metadata must not leave a
    // partially rewritten parent graph behind.
    let mut parents = Vec::new();
    for (index, row) in types.iter().enumerate() {
        let Some(class_index) = row.describes_index.checked_sub(1) else { continue };
        let parent = types[class_index as usize].parent_index;
        let descriptor_parent = if let Some(parent_index) = parent.checked_sub(1) {
            descriptor_for_class.get(parent_index as usize).copied().flatten()
                .ok_or_else(|| "class supertype has no descriptor declaration".to_owned())?
        } else {
            0
        };
        parents.push((index, descriptor_parent));
    }
    for (index, parent) in parents {
        types[index].parent_index = parent;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn class(parent_index: u16) -> TypeEntry {
        TypeEntry { name: "SameDiagnosticName".into(), parent_index, ..Default::default() }
    }

    fn descriptor(describes_index: u16) -> TypeEntry {
        TypeEntry { describes_index, ..Default::default() }
    }

    #[test]
    fn forward_declared_parent_uses_structural_descriptor_index() {
        let mut types = vec![class(2), class(0), descriptor(1), descriptor(2)];
        link_supertypes(&mut types).unwrap();
        assert_eq!(types[2].parent_index, 4);
        assert_eq!(types[3].parent_index, 0);
        assert_eq!(types[0].parent_index, 2);
    }

    #[test]
    fn descriptor_chain_preserves_each_direct_class_parent() {
        let mut types = vec![class(0), class(1), class(2), descriptor(3), descriptor(1), descriptor(2)];
        link_supertypes(&mut types).unwrap();
        assert_eq!(types[3].parent_index, 6);
        assert_eq!(types[4].parent_index, 0);
        assert_eq!(types[5].parent_index, 5);
    }

    #[test]
    fn invalid_parent_does_not_partially_publish_descriptor_links() {
        let mut types = vec![class(2), class(0), descriptor(1)];
        types[2].parent_index = 19;
        assert!(link_supertypes(&mut types).is_err());
        assert_eq!(types[2].parent_index, 19);
    }
}
