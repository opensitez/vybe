//! Unified WASM type system — GC structs + Component Model resources in one.
//!
//! Every type (user class, host type, external WASM resource) is a TypeDef:
//! - **GC struct**: indexed fields for fast access, vtable for shared methods
//! - **Component Model resource**: constructor, methods, destructor, lifecycle
//! - **Cross-language**: same dispatch path for host (Rust) and bytecode (VB/JS/C#) methods
//!
//! Objects carry a `type_id` (index into TypeRegistry).
//! Method dispatch: vtable on object's type → parent chain → Object(0).
//! Field access: indexed (`fields[i]`) for typed objects, HashMap for dynamic overflow.

use std::collections::HashMap;

#[path = "subtype_walk.rs"]
mod subtype_walk;

#[cfg(test)]
#[path = "subtype_walk_tests.rs"]
mod subtype_walk_tests;

#[cfg(test)]
#[path = "type_metadata_reserve_tests.rs"]
mod type_metadata_reserve_tests;

/// A method entry in a type's vtable.
#[derive(Debug, Clone)]
pub enum Method {
    /// Host function by index in VM's host_fns table.
    HostFn(usize),
    /// Bytecode function chunk index.
    ChunkFn(usize),
}

/// A field definition in a GC struct type.
#[derive(Debug, Clone)]
pub struct FieldDef {
    /// Field name (lowercase, for name→index resolution)
    pub name: String,
    /// Index in Object.fields[] for fast access
    pub index: usize,
    /// Property descriptor per ECMA-262 §6.2.4 (writable, enumerable, configurable).
    /// Defaults to standard() if not specified.
    pub descriptor: crate::chunk::PropertyDescriptor,
}

/// Unified type definition — covers GC structs, resources, user classes, host types.
///
/// Examples:
/// - `Button`: host type with HostFn methods, GC fields for properties
/// - `MyClass` (VB): user type with ChunkFn methods, GC fields from Dim declarations
/// - `Animal` (JS): user type with ChunkFn methods, GC fields from constructor
/// - `FileHandle` (WASI): resource with constructor + destructor + lifecycle
/// - `Color` (enum): constants only, no methods
#[derive(Debug, Clone)]
pub struct TypeDef {
    /// Type name (e.g. "List", "Button", "Animal", "FileHandle")
    pub name: String,
    pub native_identity: Option<crate::native_gc::TypeIdentity>,
    /// WASM GC composite shape (struct / array / func). Copied from the
    /// `TypeEntry` at load. `Array` makes `array.get`/`set`/`copy` on an
    /// instance apply spec trapping semantics; `Struct` (the default) is the
    /// class/struct shape used by every non-GC-array type.
    pub kind: crate::chunk::CompositeKind,
    /// Parent type index (for inheritance). None = no parent (inherits from Object/0).
    pub parent: Option<usize>,

    // -- GC struct layout --
    /// Named fields with fixed indices. `obj.fields[field.index]` for fast access.
    pub field_defs: Vec<FieldDef>,
    /// Field name → index lookup (cached from field_defs)
    pub field_map: HashMap<String, usize>,

    // -- Vtable (shared across all instances) --
    /// Method table: lowercase method name → Method (HostFn or ChunkFn)
    pub methods: HashMap<String, Method>,
    /// Constructor method (called by `New TypeName()`)
    pub constructor: Option<Method>,

    // -- Component Model resource lifecycle --
    /// Destructor — called when resource is dropped. None for non-resource types.
    pub destructor: Option<Method>,
    /// Whether this type is a managed resource (has lifecycle tracking).
    pub is_resource: bool,

    // -- Enum constants --
    pub constants: HashMap<String, i64>,

    // -- Interop --
    /// Which methods are exported cross-module (Component Model interface).
    /// Empty = all methods visible (default for user classes).
    pub exports: Vec<String>,

    // -- Interface support --
    /// Whether this type is an interface (not a concrete class).
    pub is_interface: bool,
    /// Interfaces this type implements (type_ids of interface types).
    pub implements: Vec<usize>,
    /// Required method signatures for interfaces (method_name → param_count).
    /// Only populated for interface types.
    pub required_methods: Vec<(String, u8)>,

    // -- Shared-Everything Threads --
    /// Whether instances of this type can be shared across threads.
    /// Shared types use atomic field access (shared_struct_get/set).
    pub shared: bool,

    // -- Component Model Interface Binding --
    /// Interface this type belongs to (e.g., "my:api/types").
    /// Used for cross-component type resolution.
    pub interface: Option<String>,
    /// Source component name (for tracking origin during linking).
    pub source_component: Option<String>,

    // -- WASM Annotations (Stage 2 Proposal) --
    /// Field property descriptors (WASM Annotations proposal @ecma262 namespace).
    /// Maps field_name → descriptor. Fields without entries default to PropertyDescriptor::standard().
    pub field_descriptors: HashMap<String, crate::chunk::PropertyDescriptor>,
}

#[cfg(test)]
mod field_capacity_tests {
    use super::*;

    #[test]
    fn known_field_batches_keep_storage_indices_and_descriptors() {
        for count in [0usize, 8, 32, 128, 512] {
            let mut ty = TypeDef::new("reserved_fields");
            ty.reserve_fields(count);
            let allocation = ty.field_defs.as_ptr();
            let map_capacity = ty.field_map.capacity();
            for index in 0..count {
                let mut descriptor = crate::chunk::PropertyDescriptor::standard();
                descriptor.writable = index % 2 == 0;
                let name = format!("field_{index}");
                assert_eq!(ty.add_field_with_descriptor(&name, descriptor), index);
                assert_eq!(ty.field_index(&name), Some(index));
                assert_eq!(ty.field_defs[index].descriptor.writable, index % 2 == 0);
                assert_eq!(ty.field_defs.as_ptr(), allocation);
                assert_eq!(ty.field_map.capacity(), map_capacity);
            }
            assert_eq!(ty.field_count(), count);
            assert_eq!(ty.field_index("absent"), None);
        }
    }

    #[test]
    fn equivalent_native_groups_reuse_the_registry_identity() {
        use crate::native_gc::{CompositeType, RecGroup, SubType};
        let shape = SubType {
            final_: true, shared: false, supertypes: vec![],
            descriptor: None, describes: None,
            composite: CompositeType::Func { params: vec![], results: vec![] },
        };
        let a = RecGroup::new(vec![shape.clone()], vec![]).identity(0).unwrap();
        let b = RecGroup::new(vec![shape], vec![]).identity(0).unwrap();
        let mut registry = TypeRegistry::new();
        for identity in [&a, &b] {
            registry.load_type_table(&[crate::chunk::TypeEntry {
                name: "diagnostic".into(), native_identity: Some(identity.clone()),
                kind: crate::chunk::CompositeKind::Func, ..Default::default()
            }]);
        }
        assert!(std::ptr::eq(
            registry.interned_native_identity(&a).unwrap(),
            registry.interned_native_identity(&b).unwrap(),
        ));
    }
}

impl TypeDef {
    /// Reserve the known field batch for a newly loaded module type.
    fn reserve_fields(&mut self, count: usize) {
        self.field_defs.reserve(count);
        self.field_map.reserve(count);
    }

    fn reserve_metadata(&mut self, entry: &crate::chunk::TypeEntry) {
        // Tiny batches keep their existing first-insert allocation path.
        // Descriptor keys are already unique in the input map. Required
        // signatures append every entry, including repeated method names.
        if entry.field_descriptors.len() > 3 {
            self.field_descriptors.reserve(entry.field_descriptors.len());
        }
        if entry.is_interface && entry.methods.len() > 4 {
            self.required_methods.reserve(entry.methods.len());
        }
        // Do not reserve methods by raw input length: overloads and duplicate
        // names can collapse to a much smaller final method map.
    }

    pub fn new(name: &str) -> Self {
        TypeDef {
            name: name.into(),
            native_identity: None,
            kind: crate::chunk::CompositeKind::Struct,
            parent: None,
            field_defs: Vec::new(),
            field_map: HashMap::new(),
            methods: HashMap::new(),
            constructor: None,
            destructor: None,
            is_resource: false,
            constants: HashMap::new(),
            exports: Vec::new(),
            is_interface: false,
            implements: Vec::new(),
            required_methods: Vec::new(),
            shared: false,
            interface: None,
            source_component: None,
            field_descriptors: HashMap::new(),
        }
    }

    pub fn with_parent(mut self, parent_id: usize) -> Self {
        self.parent = Some(parent_id);
        self
    }

    /// Add a named field. Returns the field index.
    pub fn add_field(&mut self, name: &str) -> usize {
        self.add_field_with_descriptor(name, crate::chunk::PropertyDescriptor::standard())
    }

    /// Add a named field with property descriptor (ECMA-262 §6.2.4).
    /// Returns the field index.
    pub fn add_field_with_descriptor(
        &mut self,
        name: &str,
        descriptor: crate::chunk::PropertyDescriptor,
    ) -> usize {
        let idx = self.field_defs.len();
        // The VM matches field names EXACTLY. Case-insensitivity is entirely a
        // compile-time concern: the compiler canonicalises names before they
        // reach here (case-sensitive langs keep case; VB/Pascal/COBOL resolve
        // to the declared name). The VM never folds case.
        self.field_defs.push(FieldDef {
            name: name.to_string(),
            index: idx,
            descriptor,
        });
        self.field_map.insert(name.to_string(), idx);
        idx
    }

    /// Get the index of a field by name (exact match), or None.
    pub fn field_index(&self, name: &str) -> Option<usize> {
        self.field_map.get(name).copied()
    }

    /// Number of indexed fields (for Object.fields pre-allocation).
    pub fn field_count(&self) -> usize {
        self.field_defs.len()
    }

    /// Whether this defined type is a WASM GC `(array …)` composite. Instances
    /// carrying this type's id apply spec trapping `array.get`/`set`/`copy`.
    pub fn is_array(&self) -> bool {
        matches!(self.kind, crate::chunk::CompositeKind::Array)
    }

    pub fn method(mut self, name: &str, m: Method) -> Self {
        self.methods.insert(name.to_string(), m);
        self
    }

    pub fn host_method(mut self, name: &str, host_fn_idx: usize) -> Self {
        self.methods
            .insert(name.to_string(), Method::HostFn(host_fn_idx));
        self
    }

    pub fn resource(mut self) -> Self {
        self.is_resource = true;
        self
    }

    pub fn shared(mut self) -> Self {
        self.shared = true;
        self
    }

    pub fn as_interface(mut self) -> Self {
        self.is_interface = true;
        self
    }

    pub fn with_required_method(mut self, name: &str, param_count: u8) -> Self {
        self.required_methods.push((name.to_string(), param_count));
        self
    }

    pub fn with_interface(mut self, iface: &str) -> Self {
        self.interface = Some(iface.to_string());
        self
    }

    /// Backward compat: old code used `fields: Vec<String>` for field names.
    /// This preserves that interface.
    pub fn fields(&self) -> Vec<String> {
        self.field_defs.iter().map(|f| f.name.clone()).collect()
    }

    /// Get property descriptor for a field (WASM Annotations proposal @ecma262 namespace).
    /// Returns the descriptor or PropertyDescriptor::standard() if not specified.
    pub fn get_field_descriptor(&self, field_name: &str) -> crate::chunk::PropertyDescriptor {
        self.field_descriptors
            .get(field_name)
            .cloned()
            .unwrap_or_else(crate::chunk::PropertyDescriptor::standard)
    }

    /// Check if a field is non-enumerable.
    pub fn is_non_enumerable(&self, field_name: &str) -> bool {
        !self.get_field_descriptor(field_name).enumerable
    }

    /// Check if a field is non-writable.
    pub fn is_non_writable(&self, field_name: &str) -> bool {
        !self.get_field_descriptor(field_name).writable
    }

    /// Check if a field is non-configurable.
    pub fn is_non_configurable(&self, field_name: &str) -> bool {
        !self.get_field_descriptor(field_name).configurable
    }
}

/// Type registry — stores all type definitions.
/// Index 0 is reserved for "Object" (universal base type).
#[derive(Debug, Clone)]
pub struct TypeRegistry {
    pub types: Vec<TypeDef>,
    /// Name → type_id lookup
    name_map: HashMap<String, usize>,
    native_types: HashMap<crate::native_gc::TypeIdentity, usize>,
}

impl TypeRegistry {
    pub fn new() -> Self {
        let mut reg = TypeRegistry {
            types: Vec::new(),
            name_map: HashMap::new(),
            native_types: HashMap::new(),
        };
        // Type 0: Object (universal base)
        reg.register(TypeDef::new("Object"));
        reg
    }

    /// Register a new type, returns its type_id.
    pub fn register(&mut self, typedef: TypeDef) -> usize {
        let id = self.types.len();
        self.name_map.insert(typedef.name.to_lowercase(), id);
        self.types.push(typedef);
        id
    }

    /// Get type_id by name. Case-insensitive: the map is keyed with lowercase
    /// (register() always lowercases), so lookup normalises to lowercase too.
    pub fn get_id(&self, name: &str) -> Option<usize> {
        // These names are already the exact key register() stores. Borrow
        // them instead of allocating a lowercase copy on every lookup.
        // Non-ASCII names retain str::to_lowercase's Unicode semantics.
        if name.bytes().all(|byte| byte.is_ascii() && !byte.is_ascii_uppercase()) {
            return self.name_map.get(name).copied();
        }
        self.name_map.get(&name.to_lowercase()).copied()
    }

    /// Lookup a key already normalized with `str::to_lowercase`.
    /// General callers must use `get_id` to preserve case-insensitive lookup.
    #[inline]
    pub(crate) fn get_lowercase_id(&self, name: &str) -> Option<usize> {
        self.name_map.get(name).copied()
    }

    pub fn get_entry_id(&self, entry: &super::chunk::TypeEntry) -> Option<usize> {
        match &entry.native_identity {
            Some(identity) => self.native_types.get(identity).copied(),
            None => self.get_id(&entry.name),
        }
    }

    pub(crate) fn get_native_id(&self, identity: &crate::native_gc::TypeIdentity) -> Option<usize> {
        self.native_types.get(identity).copied()
    }

    pub(crate) fn interned_native_identity(&self, identity: &crate::native_gc::TypeIdentity)
        -> Option<&crate::native_gc::TypeIdentity> {
        self.native_types.get_key_value(identity).map(|(identity, _)| identity)
    }

    /// Get a type definition by id.
    pub fn get(&self, type_id: usize) -> Option<&TypeDef> {
        self.types.get(type_id)
    }

    /// Get a mutable type definition by id.
    pub fn get_mut(&mut self, type_id: usize) -> Option<&mut TypeDef> {
        self.types.get_mut(type_id)
    }

    /// Resolve a field index on a type (walks parent chain for inherited fields).
    pub fn resolve_field(&self, type_id: usize, field_name: &str) -> Option<(usize, usize)> {
        let mut tid = type_id;
        loop {
            if let Some(typedef) = self.types.get(tid) {
                if let Some(&idx) = typedef.field_map.get(field_name) {
                    return Some((tid, idx));
                }
                if let Some(parent) = typedef.parent {
                    tid = parent;
                } else {
                    return None;
                }
            } else {
                return None;
            }
        }
    }

    /// Look up a method on a type, walking up the inheritance chain.
    pub fn resolve_method(&self, type_id: usize, method_name: &str) -> Option<&Method> {
        let mut tid = type_id;
        loop {
            if let Some(typedef) = self.types.get(tid) {
                if let Some(m) = typedef.methods.get(method_name) {
                    return Some(m);
                }
                if let Some(parent) = typedef.parent {
                    tid = parent;
                } else {
                    // Try universal Object type (index 0)
                    if tid != 0 {
                        return self.types[0].methods.get(method_name);
                    }
                    return None;
                }
            } else {
                return None;
            }
        }
    }

    /// Add a method to an existing type.
    pub fn add_method(&mut self, type_id: usize, method_name: &str, m: Method) {
        if let Some(typedef) = self.types.get_mut(type_id) {
            typedef.methods.insert(method_name.to_string(), m);
        }
    }

    /// Add a host method by index.
    pub fn add_host_method(&mut self, type_id: usize, method_name: &str, host_fn_idx: usize) {
        self.add_method(type_id, method_name, Method::HostFn(host_fn_idx));
    }

    /// Add a field to an existing type.
    pub fn add_field(&mut self, type_id: usize, name: &str) -> Option<usize> {
        self.types.get_mut(type_id).map(|td| td.add_field(name))
    }

    /// Set constructor for a type.
    pub fn set_constructor(&mut self, type_id: usize, m: Method) {
        if let Some(typedef) = self.types.get_mut(type_id) {
            typedef.constructor = Some(m);
        }
    }

    /// Set destructor for a resource type.
    pub fn set_destructor(&mut self, type_id: usize, m: Method) {
        if let Some(typedef) = self.types.get_mut(type_id) {
            typedef.destructor = Some(m);
            typedef.is_resource = true;
        }
    }

    /// Add a constant to a type (for enums).
    pub fn add_constant(&mut self, type_id: usize, name: &str, value: i64) {
        if let Some(typedef) = self.types.get_mut(type_id) {
            typedef.constants.insert(name.to_string(), value);
        }
    }

    /// Look up a constructor for a type.
    pub fn get_constructor(&self, type_id: usize) -> Option<&Method> {
        self.types.get(type_id)?.constructor.as_ref()
    }

    /// Look up a constant on a type (for enums).
    pub fn get_constant(&self, type_id: usize, name: &str) -> Option<i64> {
        self.types.get(type_id)?.constants.get(name).copied()
    }

    /// Export a type: mark it with an interface name so other components can import it.
    pub fn export_type(&mut self, type_id: usize, interface: &str, component: &str) {
        if let Some(td) = self.types.get_mut(type_id) {
            td.interface = Some(interface.to_string());
            td.source_component = Some(component.to_string());
        }
    }

    /// Get all types exported by a given component.
    pub fn get_component_exports(&self, component: &str) -> Vec<(usize, &TypeDef)> {
        self.types
            .iter()
            .enumerate()
            .filter(|(_, td)| td.source_component.as_deref() == Some(component))
            .collect()
    }

    /// Merge a type from another registry (for cross-component type sharing).
    /// If the type already exists, merges fields and methods.
    /// Returns the (possibly new) type_id in this registry.
    pub fn import_type(&mut self, source: &TypeDef) -> usize {
        if let Some(existing_id) = self.get_id(&source.name) {
            // Merge: add any new fields and methods
            let td = &mut self.types[existing_id];
            for fd in &source.field_defs {
                if td.field_index(&fd.name).is_none() {
                    td.add_field(&fd.name);
                }
            }
            for (name, method) in &source.methods {
                if !td.methods.contains_key(name) {
                    td.methods.insert(name.clone(), method.clone());
                }
            }
            if td.interface.is_none() {
                td.interface = source.interface.clone();
            }
            existing_id
        } else {
            // Clone the type into this registry
            let mut td = source.clone();
            // Remap parent if needed
            if let Some(parent_id) = source.parent {
                // If parent exists in source by id, try to find by name
                // For now, keep parent as-is (works if registries are compatible)
                td.parent = Some(parent_id);
            }
            self.register(td)
        }
    }

    /// Check if type_id is a subtype of target_id.
    /// Walks parent and interface edges, including inherited interfaces.
    pub fn is_subtype(&self, type_id: usize, target_id: usize) -> bool {
        if type_id == target_id {
            return true;
        }
        let Some(root) = self.types.get(type_id) else {
            return false;
        };
        // Keep common identity and direct-edge checks allocation-free.
        if root.parent == Some(target_id) || root.implements.contains(&target_id) {
            return true;
        }
        if root.parent.is_none() && root.implements.is_empty() {
            return false;
        }
        // Ordinary parent chains need no traversal worklist or visited set.
        // Start the graph walk where interfaces first appear: earlier nodes
        // have no interface edges and their parent targets were checked.
        // A registry-sized budget also bounds malformed cyclic parent chains.
        let mut graph_start = type_id;
        if root.implements.is_empty() {
            let mut next = root.parent;
            for _ in 0..self.types.len() {
                let Some(tid) = next else { return false; };
                if tid == target_id { return true; }
                let Some(typedef) = self.types.get(tid) else { return false; };
                if !typedef.implements.is_empty() {
                    if typedef.parent == Some(target_id) || typedef.implements.contains(&target_id) {
                        return true;
                    }
                    graph_start = tid;
                    break;
                }
                next = typedef.parent;
            }
        }
        subtype_walk::graph_is_subtype(self, graph_start, target_id)
    }

    /// Add an interface implementation to a type.
    pub fn add_implements(&mut self, type_id: usize, interface_id: usize) {
        if let Some(typedef) = self.types.get_mut(type_id) {
            if !typedef.implements.contains(&interface_id) {
                typedef.implements.push(interface_id);
            }
        }
    }

    /// Register an interface type with required methods.
    pub fn register_interface(&mut self, name: &str, methods: &[(&str, u8)]) -> usize {
        let mut td = TypeDef::new(name).as_interface();
        for (method_name, param_count) in methods {
            td.required_methods
                .push((method_name.to_string(), *param_count));
        }
        self.register(td)
    }

    /// Check if a type satisfies an interface (has all required methods).
    pub fn satisfies_interface(&self, type_id: usize, interface_id: usize) -> bool {
        let iface = match self.types.get(interface_id) {
            Some(td) if td.is_interface => td,
            _ => return false,
        };
        for (method_name, _) in &iface.required_methods {
            if self.resolve_method(type_id, method_name).is_none() {
                return false;
            }
        }
        true
    }

    /// Get the constructor for a type, walking parent chain if needed.
    pub fn resolve_constructor(&self, type_id: usize) -> Option<&Method> {
        let mut tid = type_id;
        loop {
            if let Some(typedef) = self.types.get(tid) {
                if typedef.constructor.is_some() {
                    return typedef.constructor.as_ref();
                }
                if let Some(parent) = typedef.parent {
                    tid = parent;
                } else {
                    return None;
                }
            } else {
                return None;
            }
        }
    }

    /// Load type entries from a compiled chunk's type table.
    /// Called by VM.run() before execution. Registers user-defined types
    /// and adds their ChunkFn methods to the vtable.
    pub fn load_type_table(&mut self, types: &[super::chunk::TypeEntry]) {
        // First pass: register every type, and record where each entry of THIS
        // table landed. `local_ids[i]` is the registry id of `types[i]`, which
        // is what turns the declared `sub $i` links below into id links
        // without going near a name. An entry whose name is already registered
        // — a host builtin the module only DECLARED so it could name it —
        // keeps its existing id, which is exactly the binding an imported type
        // needs.
        let mut local_ids: Vec<usize> = Vec::with_capacity(types.len());
        for entry in types {
            let id = match self.get_entry_id(entry) {
                Some(existing) => existing,
                None => {
                    let mut td = TypeDef::new(&entry.name);
                    td.native_identity = entry.native_identity.clone();
                    if let Some(identity) = &td.native_identity {
                        td.shared = identity.definition().shared;
                    }
                    td.reserve_fields(entry.fields.len());
                    td.reserve_metadata(entry);
                    // Carry the WASM GC composite shape so instances stamped
                    // with this type id trap on `array.*` when it is an
                    // `(array …)`.
                    td.kind = entry.kind;
                    if entry.is_interface {
                        td.is_interface = true;
                    }
                    if let Some(identity) = &entry.native_identity {
                        let id = self.types.len();
                        self.types.push(td);
                        self.native_types.insert(identity.clone(), id);
                        id
                    } else {
                        self.register(td)
                    }
                }
            };
            local_ids.push(id);
        }

        // Second pass over the supertype links, now that every id is known.
        // Forward references are free: a class may declare a supertype that
        // appears later in the same table, exactly as a WASM rec group allows.
        for (i, entry) in types.iter().enumerate() {
            if entry.parent_index == 0 {
                continue;
            }
            if let Some(&parent_id) = local_ids.get(entry.parent_index as usize - 1) {
                if parent_id != local_ids[i] {
                    self.types[local_ids[i]].parent = Some(parent_id);
                }
            }
        }

        // Second pass: add fields, methods, interface implementations, constructors
        // Registration has already resolved every entry, including duplicate
        // and imported names. Reuse those IDs rather than lowercasing and
        // hashing every name again. Fields and methods do not change name_map.
        for (entry, &type_id) in types.iter().zip(&local_ids) {

            // Add fields
            {
                let typedef = &mut self.types[type_id];
                for field_name in &entry.fields {
                    if typedef.field_index(field_name).is_none() {
                        typedef.add_field(field_name);
                    }
                }

                // Copy field descriptors from TypeEntry (WASM Annotations proposal @ecma262)
                for (field_name, descriptor) in &entry.field_descriptors {
                    typedef
                        .field_descriptors
                        .insert(field_name.clone(), descriptor.clone());
                }

                // Mark as interface if flagged
                if entry.is_interface {
                    typedef.is_interface = true;
                    // For interfaces, methods are required method signatures
                    for (method_name, _) in &entry.methods {
                        typedef.required_methods.push((method_name.clone(), 0));
                    }
                }

                // Add vtable methods (ChunkFn)
                for (method_name, chunk_idx) in &entry.methods {
                    typedef
                        .methods
                        .insert(method_name.clone(), Method::ChunkFn(*chunk_idx));
                }

                // Set constructor
                if let Some(ctor_idx) = entry.constructor_chunk {
                    typedef.constructor = Some(Method::ChunkFn(ctor_idx));
                }
            }

            // Resolve interface implementations by INDEX, through the same
            // `local_ids` map the supertype links use.
            for &iface_index in &entry.implements {
                if iface_index == 0 {
                    continue;
                }
                if let Some(&iface_id) = local_ids.get(iface_index as usize - 1) {
                    if iface_id != type_id {
                        self.add_implements(type_id, iface_id);
                    }
                }
            }
        }
    }
}

/// Resource table — tracks live resource handles for ownership/borrowing.
/// Used for Component Model resource types (files, sockets, DB connections).
#[derive(Debug, Default)]
pub struct ResourceTable {
    entries: HashMap<u32, ResourceEntry>,
    next_handle: u32,
}

#[derive(Debug, Clone)]
struct ResourceEntry {
    type_id: usize,
    data: super::Value,
    borrow_count: u32,
    dropped: bool,
}

impl ResourceTable {
    pub fn new() -> Self {
        ResourceTable {
            entries: HashMap::new(),
            next_handle: 1,
        }
    }

    pub fn create(&mut self, type_id: usize, data: super::Value) -> u32 {
        let handle = self.next_handle;
        self.next_handle += 1;
        self.entries.insert(
            handle,
            ResourceEntry {
                type_id,
                data,
                borrow_count: 0,
                dropped: false,
            },
        );
        handle
    }

    pub fn borrow(&mut self, handle: u32) -> Option<super::Value> {
        let entry = self.entries.get_mut(&handle)?;
        if entry.dropped {
            return None;
        }
        entry.borrow_count += 1;
        Some(entry.data.clone())
    }

    pub fn release_borrow(&mut self, handle: u32) {
        if let Some(entry) = self.entries.get_mut(&handle) {
            if entry.borrow_count > 0 {
                entry.borrow_count -= 1;
            }
        }
    }

    pub fn drop_resource(&mut self, handle: u32) -> Result<super::Value, String> {
        let entry = self
            .entries
            .get_mut(&handle)
            .ok_or_else(|| format!("Invalid resource handle: {}", handle))?;
        if entry.dropped {
            return Err(format!("Resource {} already dropped", handle));
        }
        if entry.borrow_count > 0 {
            return Err(format!(
                "Cannot drop resource {}: {} active borrows",
                handle, entry.borrow_count
            ));
        }
        entry.dropped = true;
        Ok(entry.data.clone())
    }

    pub fn type_id(&self, handle: u32) -> Option<usize> {
        self.entries.get(&handle).map(|e| e.type_id)
    }

    pub fn is_valid(&self, handle: u32) -> bool {
        self.entries
            .get(&handle)
            .map(|e| !e.dropped)
            .unwrap_or(false)
    }
}

#[cfg(test)]
mod normalized_type_lookup_tests {
    #[test]
    fn normalized_lookup_preserves_general_lookup_results() {
        let mut vm = crate::VM::new();
        let registry = &mut vm.type_registry;
        // Isolate name-index lookup; no type definitions are needed here.
        for (index, name) in ["ClassName", "Namespace.Class", "İ", "Σςσ", "Straße", "中文", "A\0B"].into_iter().enumerate() {
            registry.name_map.insert(name.to_lowercase(), index + 100);
        }
        for name in ["ClassName", "CLASSNAME", "Namespace.Class", "namespace.class", "İ", "i\u{307}", "Σςσ", "Straße", "中文", "A\0B", "missing", ""] {
            let canonical = name.to_lowercase();
            assert_eq!(registry.get_lowercase_id(&canonical), registry.get_id(name));
            assert_eq!(registry.get_lowercase_id(&canonical), registry.get_id(&canonical));
        }
    }
}
