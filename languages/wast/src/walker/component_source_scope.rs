//! Lexical declaration identity and separate component/core binary namespaces.
use super::{Pair, Rule};
use std::collections::HashMap;
use std::sync::Arc;
use vybe_ast::canon::{ComponentItemKind as Kind, ValSpec};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub(super) enum Namespace {
    Module,
    Func,
    Value,
    Type,
    Component,
    Instance,
    CoreFunc,
    CoreTable,
    CoreMemory,
    CoreGlobal,
    CoreTag,
    CoreType,
    CoreInstance,
}

impl Namespace {
    pub fn slot(self) -> usize {
        self as usize
    }
}

impl From<Kind> for Namespace {
    fn from(kind: Kind) -> Self {
        match kind {
            Kind::Module => Self::Module,
            Kind::Func => Self::Func,
            Kind::Value => Self::Value,
            Kind::Type => Self::Type,
            Kind::Component => Self::Component,
            Kind::Instance => Self::Instance,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub(super) struct DeclarationId {
    pub scope: (usize, usize),
    pub namespace: Namespace,
    pub ordinal: u32,
}

#[derive(Clone, Debug)]
pub(super) struct Binding {
    pub declaration: DeclarationId,
    pub origin: DeclarationId,
    pub source: (usize, usize),
}

#[derive(Clone, Debug)]
pub(super) struct TypeSource {
    pub declaration: DeclarationId,
    pub shape: SourceShape,
}

#[derive(Clone, Debug)]
pub(super) enum SourceShape {
    Value(ValSpec),
    Interface(SourceInterface),
    Opaque,
}

#[derive(Clone, Debug, Default)]
pub(super) struct SourceInterface {
    pub is_component: bool,
    pub exports: HashMap<String, SourceItem>,
}

#[derive(Clone, Debug)]
pub(super) struct SourceItem {
    pub namespace: Namespace,
    pub origin: DeclarationId,
    pub projection: Option<Arc<TypeSource>>,
}

#[derive(Clone, Default)]
pub(super) struct SourceScope {
    pub label: Option<String>,
    pub names: [HashMap<String, u32>; 13],
    pub counts: [u32; 13],
    pub bindings: [Vec<Binding>; 13],
    // Keys are introduced declarations, never external names or structural types.
    projections: HashMap<DeclarationId, Arc<TypeSource>>,
    pub exports: HashMap<String, SourceItem>,
    pub core_context: Vec<String>,
    pub core_definitions: usize,
    id: (usize, usize),
    pub site: (usize, usize),
}

impl SourceScope {
    pub fn new(source: &Pair<'_, Rule>) -> Self {
        let span = source.as_span();
        Self {
            id: (span.start(), span.end()),
            site: (span.start(), span.end()),
            label: source
                .clone()
                .into_inner()
                .find(|p| p.as_rule() == Rule::id)
                .map(|p| p.as_str()[1..].to_owned()),
            ..Self::default()
        }
    }

    pub fn add(&mut self, kind: Kind, name: Option<String>) -> Result<u32, String> {
        self.define(kind.into(), name, None)
    }

    pub fn define(
        &mut self,
        namespace: Namespace,
        name: Option<String>,
        origin: Option<DeclarationId>,
    ) -> Result<u32, String> {
        let slot = namespace.slot();
        let index = self.counts[slot];
        self.counts[slot] = index.checked_add(1).ok_or("component index overflow")?;
        if let Some(name) = name {
            if self.names[slot].insert(name.clone(), index).is_some() {
                return Err(format!("duplicate {namespace:?} name ${name}"));
            }
        }
        let declaration = DeclarationId {
            scope: self.id,
            namespace,
            ordinal: index,
        };
        self.bindings[slot].push(Binding {
            declaration,
            origin: origin.unwrap_or(declaration),
            source: self.site,
        });
        Ok(index)
    }

    fn declaration(&self, namespace: Namespace, index: u32) -> Result<DeclarationId, String> {
        self.bindings[namespace.slot()]
            .get(index as usize)
            .map(|b| b.declaration)
            .ok_or_else(|| format!("{namespace:?} index out of bounds"))
    }

    pub fn projection(&self, namespace: Namespace, index: u32) -> Option<Arc<TypeSource>> {
        self.declaration(namespace, index)
            .ok()
            .and_then(|id| self.projections.get(&id).cloned())
    }

    pub fn project(
        &mut self,
        namespace: Namespace,
        index: u32,
        source: Option<Arc<TypeSource>>,
    ) -> Result<(), String> {
        let id = self.declaration(namespace, index)?;
        if let Some(source) = source {
            self.projections.insert(id, source);
        } else {
            self.projections.remove(&id);
        }
        Ok(())
    }

    pub fn define_shape(
        &mut self,
        namespace: Namespace,
        index: u32,
        shape: SourceShape,
    ) -> Result<(), String> {
        let declaration = self.declaration(namespace, index)?;
        self.project(
            namespace,
            index,
            Some(Arc::new(TypeSource { declaration, shape })),
        )
    }

    pub fn item(&self, namespace: Namespace, index: u32) -> Result<SourceItem, String> {
        Ok(SourceItem {
            namespace,
            origin: self.origin(namespace, index)?,
            projection: self.projection(namespace, index),
        })
    }

    pub fn alias(
        &mut self,
        namespace: Namespace,
        name: Option<String>,
        item: &SourceItem,
    ) -> Result<u32, String> {
        if namespace != item.namespace {
            return Err("source export has the wrong namespace".into());
        }
        let index = self.define(namespace, name, Some(item.origin))?;
        self.project(namespace, index, item.projection.clone())?;
        Ok(index)
    }

    pub fn publish(
        &mut self,
        name: String,
        namespace: Namespace,
        index: u32,
    ) -> Result<(), String> {
        self.exports.insert(name, self.item(namespace, index)?);
        Ok(())
    }

    pub fn interface(&self, is_component: bool) -> SourceInterface {
        SourceInterface {
            is_component,
            exports: self.exports.clone(),
        }
    }

    pub fn export(
        &self,
        namespace: Namespace,
        index: u32,
        name: &str,
    ) -> Result<SourceItem, String> {
        let source = self
            .projection(namespace, index)
            .ok_or("source interface projection is unavailable")?;
        let SourceShape::Interface(interface) = &source.shape else {
            return Err("source item is not an interface".into());
        };
        if (namespace == Namespace::Component) != interface.is_component {
            return Err("source interface has the wrong namespace".into());
        }
        interface
            .exports
            .get(name)
            .cloned()
            .ok_or_else(|| format!("source interface has no export `{name}`"))
    }

    pub fn origin(&self, namespace: Namespace, index: u32) -> Result<DeclarationId, String> {
        self.bindings[namespace.slot()]
            .get(index as usize)
            .map(|b| b.origin)
            .ok_or_else(|| format!("{namespace:?} index out of bounds"))
    }

    pub fn resolve(&self, namespace: Namespace, pair: &Pair<'_, Rule>) -> Result<u32, String> {
        let index = resolve_idx(pair, "component item", &self.names[namespace.slot()])?;
        if index >= self.counts[namespace.slot()] {
            return Err(format!("{namespace:?} index out of bounds"));
        }
        Ok(index)
    }

    pub fn index(&self, kind: Kind, pair: &Pair<'_, Rule>) -> Result<u32, String> {
        self.resolve(kind.into(), pair)
    }

    pub fn check(&self) -> Result<(), String> {
        for (id, source) in &self.projections {
            if self.declaration(id.namespace, id.ordinal)? != *id {
                return Err("source projection is not anchored to its declaration".into());
            }
            if matches!(source.shape, SourceShape::Value(_))
                && source.declaration.namespace != Namespace::Type
            {
                return Err("source value shape has no type declaration".into());
            }
        }
        for (slot, bindings) in self.bindings.iter().enumerate() {
            if bindings.len() != self.counts[slot] as usize {
                return Err("source namespace slot count mismatch".into());
            }
            for (ordinal, binding) in bindings.iter().enumerate() {
                if binding.declaration.scope != self.id
                    || binding.declaration.namespace.slot() != slot
                    || binding.declaration.ordinal as usize != ordinal
                    || binding.source.0 < self.id.0
                    || binding.source.1 > self.id.1
                {
                    return Err("source declaration identity mismatch".into());
                }
            }
        }
        Ok(())
    }
}

pub(super) fn resolve_idx(
    pair: &Pair<'_, Rule>,
    space: &str,
    names: &HashMap<String, u32>,
) -> Result<u32, String> {
    let raw = pair.as_str().trim();
    match raw.strip_prefix('$') {
        Some(name) => names
            .get(name)
            .copied()
            .ok_or_else(|| format!("canon: `${name}` is not bound in the {space} index space")),
        None => raw
            .replace('_', "")
            .parse::<u32>()
            .map_err(|_| format!("canon: `{raw}` is not a {space} index")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn component(source: &str) -> Pair<'_, Rule> {
        crate::WastParser::parse(Rule::component, source)
            .unwrap()
            .next()
            .unwrap()
    }

    #[test]
    fn every_namespace_is_independent_and_case_sensitive() {
        let source = component("(component $Root)");
        let mut scope = SourceScope::new(&source);
        for namespace in [
            Namespace::Module,
            Namespace::Func,
            Namespace::Value,
            Namespace::Type,
            Namespace::Component,
            Namespace::Instance,
            Namespace::CoreFunc,
            Namespace::CoreTable,
            Namespace::CoreMemory,
            Namespace::CoreGlobal,
            Namespace::CoreTag,
            Namespace::CoreType,
            Namespace::CoreInstance,
        ] {
            assert_eq!(
                scope.define(namespace, Some("Name".into()), None).unwrap(),
                0
            );
            assert_eq!(
                scope.define(namespace, Some("name".into()), None).unwrap(),
                1
            );
            let reference = crate::WastParser::parse(Rule::index, "$Name")
                .unwrap()
                .next()
                .unwrap();
            assert_eq!(scope.resolve(namespace, &reference).unwrap(), 0);
            assert_eq!(scope.origin(namespace, 0).unwrap().namespace, namespace);
        }
        let reference = crate::WastParser::parse(Rule::index, "$name")
            .unwrap()
            .next()
            .unwrap();
        assert_eq!(scope.index(Kind::Value, &reference).unwrap(), 1);
        assert_eq!(scope.label.as_deref(), Some("Root"));
        scope.core_context.push("(core type (func))".into());
        scope.core_definitions = 1;
        assert_eq!(scope.clone().core_context.len(), scope.core_definitions);
        scope.check().unwrap();
    }

    #[test]
    fn aliases_have_new_declarations_but_preserve_origin() {
        let source = component("(component)");
        let mut scope = SourceScope::new(&source);
        let original = scope.add(Kind::Value, Some("v".into())).unwrap();
        let origin = scope.origin(Namespace::Value, original).unwrap();
        let alias = scope
            .define(Namespace::Value, Some("public".into()), Some(origin))
            .unwrap();
        let a = &scope.bindings[Namespace::Value.slot()][alias as usize];
        assert_ne!(a.declaration, a.origin);
        assert_eq!(a.origin, origin);
        assert_eq!(a.declaration.ordinal, 1);
        scope.check().unwrap();
    }

    #[test]
    fn lexical_shadowing_does_not_merge_declaration_identity() {
        let root = component("(component $Root (component $Child))");
        let child = root
            .clone()
            .into_inner()
            .find(|p| p.as_rule() == Rule::component_definition)
            .unwrap()
            .into_inner()
            .next()
            .unwrap();
        let mut outer = SourceScope::new(&root);
        let mut inner = SourceScope::new(&child);
        outer.add(Kind::Type, Some("same".into())).unwrap();
        inner.add(Kind::Type, Some("same".into())).unwrap();
        assert_ne!(
            outer.origin(Namespace::Type, 0).unwrap(),
            inner.origin(Namespace::Type, 0).unwrap()
        );
        inner
            .define(Namespace::CoreType, Some("same".into()), None)
            .unwrap();
        assert_ne!(
            inner.origin(Namespace::Type, 0).unwrap(),
            inner.origin(Namespace::CoreType, 0).unwrap()
        );
        outer.check().unwrap();
        inner.check().unwrap();
    }

    #[test]
    fn same_named_equal_shapes_in_distinct_scopes_keep_distinct_sources() {
        let source = component("(component (type $A (instance)) (type $B (instance)))");
        let mut bodies = source
            .clone()
            .into_inner()
            .filter(|p| p.as_rule() == Rule::component_definition)
            .map(|p| {
                p.into_inner()
                    .next()
                    .unwrap()
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::deftype)
                    .unwrap()
                    .into_inner()
                    .next()
                    .unwrap()
            });
        let mut left = SourceScope::new(&bodies.next().unwrap());
        let mut right = SourceScope::new(&bodies.next().unwrap());
        for scope in [&mut left, &mut right] {
            let ty = scope.add(Kind::Type, Some("t".into())).unwrap();
            scope
                .define_shape(
                    Namespace::Type,
                    ty,
                    SourceShape::Value(ValSpec::Prim("u8".into())),
                )
                .unwrap();
            scope
                .publish("payload".into(), Namespace::Type, ty)
                .unwrap();
            let other = scope.add(Kind::Type, None).unwrap();
            scope
                .define_shape(Namespace::Type, other, SourceShape::Opaque)
                .unwrap();
            scope.check().unwrap();
        }
        let mut root = SourceScope::new(&source);
        for interface in [left.interface(false), right.interface(false)] {
            let instance = root.add(Kind::Instance, None).unwrap();
            root.define_shape(
                Namespace::Instance,
                instance,
                SourceShape::Interface(interface),
            )
            .unwrap();
            let item = root
                .export(Namespace::Instance, instance, "payload")
                .unwrap();
            root.alias(Namespace::Type, None, &item).unwrap();
        }
        let a = root.projection(Namespace::Type, 0).unwrap();
        let b = root.projection(Namespace::Type, 1).unwrap();
        assert_ne!(a.declaration, b.declaration);
        assert!(matches!(&a.shape,SourceShape::Value(ValSpec::Prim(t)) if t=="u8"));
        assert!(matches!(&b.shape,SourceShape::Value(ValSpec::Prim(t)) if t=="u8"));
        let repeated = root.export(Namespace::Instance, 0, "payload").unwrap();
        let index = root.alias(Namespace::Type, None, &repeated).unwrap();
        assert_eq!(index, 2);
        assert_eq!(
            root.projection(Namespace::Type, 2).unwrap().declaration,
            a.declaration
        );
        assert_ne!(
            root.bindings[Namespace::Type.slot()][0].declaration,
            root.bindings[Namespace::Type.slot()][2].declaration
        );
        assert!(root.alias(Namespace::Value, None, &repeated).is_err());
        assert!(root.export(Namespace::Instance, 0, "Payload").is_err());
        root.check().unwrap();
    }
}
