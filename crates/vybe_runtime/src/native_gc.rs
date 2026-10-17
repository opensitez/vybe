//! Structural identities for native WebAssembly recursive type groups.

use std::hash::{Hash, Hasher};
use std::sync::Arc;

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum TypeRef {
    Local(u32),
    Dependency(u32),
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum HeapType {
    Abstract { kind: u8, shared: bool },
    Concrete(TypeRef),
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum ValueType {
    Scalar(u8),
    Ref {
        nullable: bool,
        exact: bool,
        heap: HeapType,
    },
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum StorageType {
    I8,
    I16,
    Value(ValueType),
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct FieldType {
    pub storage: StorageType,
    pub mutable: bool,
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum CompositeType {
    Func {
        params: Vec<ValueType>,
        results: Vec<ValueType>,
    },
    Struct(Vec<FieldType>),
    Array(FieldType),
    Cont(TypeRef),
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct SubType {
    pub final_: bool,
    pub shared: bool,
    pub supertypes: Vec<TypeRef>,
    pub descriptor: Option<TypeRef>,
    pub describes: Option<TypeRef>,
    pub composite: CompositeType,
}

pub struct RecGroup {
    types: Vec<SubType>,
    dependencies: Vec<TypeIdentity>,
    fingerprint: u64,
}

impl std::fmt::Debug for RecGroup {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("RecGroup")
            .field("types", &self.types)
            .field("dependency_count", &self.dependencies.len())
            .field("fingerprint", &self.fingerprint)
            .finish()
    }
}

impl RecGroup {
    pub fn new(types: Vec<SubType>, dependencies: Vec<TypeIdentity>) -> Arc<Self> {
        let mut hash = std::collections::hash_map::DefaultHasher::new();
        types.hash(&mut hash);
        dependencies.hash(&mut hash);
        Arc::new(Self {
            types,
            dependencies,
            fingerprint: hash.finish(),
        })
    }

    pub fn identity(self: &Arc<Self>, position: u32) -> Option<TypeIdentity> {
        ((position as usize) < self.types.len()).then(|| TypeIdentity {
            group: Arc::clone(self),
            position,
        })
    }
}

#[derive(Clone)]
pub struct TypeIdentity {
    group: Arc<RecGroup>,
    position: u32,
}

impl std::fmt::Debug for TypeIdentity {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("TypeIdentity")
            .field("group", &self.group.fingerprint)
            .field("position", &self.position)
            .field("definition", self.definition())
            .finish()
    }
}

impl TypeIdentity {
    pub fn definition(&self) -> &SubType {
        &self.group.types[self.position as usize]
    }
}

impl PartialEq for TypeIdentity {
    fn eq(&self, other: &Self) -> bool {
        if self.position != other.position {
            return false;
        }
        if Arc::ptr_eq(&self.group, &other.group) {
            return true;
        }
        if self.group.fingerprint != other.group.fingerprint {
            return false;
        }
        if self.group.types != other.group.types
            || self.group.dependencies.len() != other.group.dependencies.len()
        {
            return false;
        }
        if self.group.dependencies.is_empty() {
            return true;
        }
        // Dependency graphs may be deep or share many ancestors. Do not
        // recurse on the native stack or revisit the same shared pair.
        let mut pending = vec![(&self.group, &other.group)];
        let mut visited = std::collections::HashSet::new();
        while let Some((a, b)) = pending.pop() {
            if Arc::ptr_eq(a, b) {
                continue;
            }
            if !visited.insert((Arc::as_ptr(a), Arc::as_ptr(b))) {
                continue;
            }
            if a.fingerprint != b.fingerprint
                || a.types != b.types
                || a.dependencies.len() != b.dependencies.len()
            {
                return false;
            }
            for (a, b) in a.dependencies.iter().zip(&b.dependencies) {
                if a.position != b.position {
                    return false;
                }
                pending.push((&a.group, &b.group));
            }
        }
        true
    }
}

impl Eq for TypeIdentity {}

impl Hash for TypeIdentity {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.group.fingerprint.hash(state);
        self.position.hash(state);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn structure(storage: StorageType) -> SubType {
        SubType {
            final_: true,
            shared: false,
            supertypes: vec![],
            descriptor: None,
            describes: None,
            composite: CompositeType::Struct(vec![FieldType {
                storage,
                mutable: false,
            }]),
        }
    }

    #[test]
    fn recursive_identity_is_structural_and_group_sensitive() {
        let recursive = structure(StorageType::Value(ValueType::Ref {
            nullable: true,
            exact: false,
            heap: HeapType::Concrete(TypeRef::Local(0)),
        }));
        let a = RecGroup::new(vec![recursive.clone()], vec![]);
        let b = RecGroup::new(vec![recursive.clone()], vec![]);
        assert_eq!(a.identity(0), b.identity(0));
        let larger = RecGroup::new(vec![recursive.clone(), recursive], vec![]);
        assert_ne!(a.identity(0), larger.identity(0));
        assert_ne!(larger.identity(0), larger.identity(1));
        assert!(a.identity(1).is_none());
    }

    #[test]
    fn hash_collisions_do_not_merge_unrelated_types() {
        let a = RecGroup::new(vec![structure(StorageType::I8)], vec![]);
        let mut b = RecGroup::new(vec![structure(StorageType::I16)], vec![]);
        Arc::get_mut(&mut b).unwrap().fingerprint = a.fingerprint;
        let mut identities = std::collections::HashMap::new();
        identities.insert(a.identity(0).unwrap(), 1);
        identities.insert(b.identity(0).unwrap(), 2);
        assert_eq!(identities.len(), 2);
    }

    #[test]
    fn deep_shared_dependencies_compare_without_recursive_walks() {
        let mut a = vec![RecGroup::new(vec![structure(StorageType::I8)], vec![])];
        let mut b = vec![RecGroup::new(vec![structure(StorageType::I8)], vec![])];
        for _ in 0..4_096 {
            let field = structure(StorageType::Value(ValueType::Ref {
                nullable: true,
                exact: false,
                heap: HeapType::Concrete(TypeRef::Dependency(0)),
            }));
            let previous_a = a.last().unwrap().identity(0).unwrap();
            let previous_b = b.last().unwrap().identity(0).unwrap();
            a.push(RecGroup::new(vec![field.clone()], vec![previous_a]));
            b.push(RecGroup::new(vec![field], vec![previous_b]));
        }
        assert_eq!(a.last().unwrap().identity(0), b.last().unwrap().identity(0));
        while a.pop().is_some() {}
        while b.pop().is_some() {}
    }
}
