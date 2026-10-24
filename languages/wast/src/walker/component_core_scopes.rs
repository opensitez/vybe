//! Source-schema slots only. Native identities and validity belong to Wasm's resolver.
use std::sync::Arc;
use vybe_ast::canon::{ComponentBinaryDef as Def, ComponentCoreSubtype, ComponentCoreType};

#[derive(Clone, Debug)]
pub(super) struct Slot {
    pub definition: Arc<ComponentCoreType>,
    pub member: Option<usize>,
}

impl Slot {
    pub fn subtype(&self) -> Option<&ComponentCoreSubtype> {
        match self.definition.as_ref() {
            ComponentCoreType::Rec(group) => group.types.get(self.member?),
            ComponentCoreType::Module(_) => None,
        }
    }
}

#[derive(Clone, Debug, Default)]
pub(super) struct Scope {
    slots: Vec<Slot>,
}

impl Scope {
    #[cfg(test)]
    pub fn len(&self) -> usize { self.slots.len() }

    pub fn slot(&self, index: u32) -> Result<&Slot, String> {
        self.slots.get(index as usize).ok_or_else(|| "core type index out of bounds".into())
    }

    pub fn define(&mut self, ty: ComponentCoreType) -> Result<Vec<Slot>, String> {
        let count = ty.type_count().ok_or("core type count exceeds u32")? as usize;
        let end = self.slots.len().checked_add(count).ok_or("core type index overflow")?;
        u32::try_from(end).map_err(|_| "core type index overflow")?;
        let module = matches!(ty, ComponentCoreType::Module(_));
        let definition = Arc::new(ty);
        let slots: Vec<_> = (0..count).map(|member| Slot {
            definition: definition.clone(), member: (!module).then_some(member),
        }).collect();
        self.slots.extend(slots.iter().cloned());
        Ok(slots)
    }

    pub fn alias(&mut self, outers: &[Scope], count: u32, index: u32) -> Result<Slot, String> {
        let source = if count == 0 { &*self } else {
            outers.get(outers.len().checked_sub(count as usize)
                .ok_or("core outer alias scope out of bounds")?)
                .ok_or("core outer alias scope out of bounds")?
        };
        let slot = source.slot(index)?.clone();
        u32::try_from(self.slots.len().checked_add(1).ok_or("core type index overflow")?)
            .map_err(|_| "core type index overflow")?;
        self.slots.push(slot.clone());
        Ok(slot)
    }

    pub fn declaration(&mut self, definition: &Def, outers: &[Scope]) -> Result<Option<Vec<Slot>>, String> {
        if let Some(ty) = definition.normalized_core_type() {
            return self.define(ty.into_owned()).map(Some);
        }
        if let Some(alias) = definition.normalized_outer_alias() {
            if alias.kind == vybe_ast::canon::ComponentOuterAliasKind::CoreType {
                return self.alias(outers, alias.count, alias.index).map(|slot| Some(vec![slot]));
            }
        }
        Ok(None)
    }
}
