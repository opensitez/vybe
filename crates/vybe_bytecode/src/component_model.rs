//! Component Model resource patterns for WASM GC integration.
//!
//! This module extends the Component Model (component.rs) with resource types,
//! which represent opaque handles to host-managed state. Resources follow the
//! WASM Component Model's `resource` pattern:
//!
//! - A resource has a constructor, methods, and a destructor (drop).
//! - Resources are owned — they have a single owner and are dropped when the owner
//!   goes out of scope (or calls `[resource-drop]`).
//! - Resources can be borrowed — `borrow<T>` gives temporary access without ownership.
//!
//! Examples:
//! - A file handle is a resource: `open()` creates it, `read()`/`write()` use it,
//!   `close()` or drop destroys it.
//! - A GUI control is a resource: the host manages the actual widget, the guest
//!   holds an opaque handle.
//! - A database connection is a resource.
//!
//! Integration with TypeRegistry:
//! - Each resource type is registered in TypeRegistry (has a type_id).
//! - Resource instances carry `type_id` on the Object.
//! - Resource methods are resolved via vtable (TypeRegistry.resolve_method).
//! - `ref_test` / `instanceof` work with resource types.

use std::collections::HashMap;

/// A resource type definition in the Component Model.
/// Resources are opaque handles to host-managed state.
#[derive(Debug, Clone)]
pub struct ResourceType {
    /// Resource type name (e.g., "FileHandle", "DbConnection", "Control")
    pub name: String,
    /// Type ID in the TypeRegistry (set after registration)
    pub type_id: usize,
    /// Constructor signature: param types → resource handle
    pub constructor: Option<ResourceMethod>,
    /// Named methods on this resource
    pub methods: Vec<ResourceMethod>,
    /// Destructor — called when the resource is dropped
    pub destructor: Option<ResourceMethod>,
    /// Whether this resource can be borrowed (vs. always owned)
    pub borrowable: bool,
}

/// A method on a resource type.
#[derive(Debug, Clone)]
pub struct ResourceMethod {
    /// Method name (e.g., "read", "write", "close")
    pub name: String,
    /// Whether this method takes `self` (instance method) or not (static method).
    pub is_static: bool,
    /// Whether this method takes `borrow<self>` (read-only) or `own<self>` (consuming).
    pub borrows_self: bool,
    /// Parameter types (excluding self)
    pub params: Vec<super::component::ValType>,
    /// Return type(s)
    pub results: Vec<super::component::ValType>,
}

/// A component import/export kind — extends the basic Component with resource support.
#[derive(Debug, Clone)]
pub enum ComponentItemKind {
    /// A function import/export
    Function(super::component::FuncSig),
    /// A resource type import/export
    Resource(ResourceType),
    /// A type alias (e.g., `type handle = u32`)
    Type(TypeAlias),
}

/// A type alias in the Component Model.
#[derive(Debug, Clone)]
pub struct TypeAlias {
    pub name: String,
    pub target: super::component::ValType,
}

/// A component import with kind information.
#[derive(Debug, Clone)]
pub struct ComponentImport {
    /// Interface name (e.g., "wasi:filesystem/types")
    pub interface: String,
    /// Item name within the interface
    pub name: String,
    /// What kind of import this is
    pub kind: ComponentItemKind,
}

/// A component export with kind information.
#[derive(Debug, Clone)]
pub struct ComponentExport {
    /// Interface name
    pub interface: String,
    /// Item name
    pub name: String,
    /// What kind of export this is
    pub kind: ComponentItemKind,
}

/// A complete component descriptor with resource support.
/// This extends `Component` (component.rs) with richer type information.
#[derive(Debug, Clone)]
pub struct ComponentDescriptor {
    /// Component name
    pub name: String,
    /// Typed imports (functions + resources + types)
    pub imports: Vec<ComponentImport>,
    /// Typed exports (functions + resources + types)
    pub exports: Vec<ComponentExport>,
    /// Resource types defined by this component
    pub resources: Vec<ResourceType>,
}

impl ComponentDescriptor {
    pub fn new(name: impl Into<String>) -> Self {
        ComponentDescriptor {
            name: name.into(),
            imports: Vec::new(),
            exports: Vec::new(),
            resources: Vec::new(),
        }
    }

    /// Add a resource type definition to this component.
    pub fn add_resource(&mut self, resource: ResourceType) {
        self.resources.push(resource);
    }

    /// Add a function import.
    pub fn add_import_fn(&mut self, interface: &str, name: &str, sig: super::component::FuncSig) {
        self.imports.push(ComponentImport {
            interface: interface.into(),
            name: name.into(),
            kind: ComponentItemKind::Function(sig),
        });
    }

    /// Add a resource import.
    pub fn add_import_resource(&mut self, interface: &str, name: &str, resource: ResourceType) {
        self.imports.push(ComponentImport {
            interface: interface.into(),
            name: name.into(),
            kind: ComponentItemKind::Resource(resource),
        });
    }

    /// Add a function export.
    pub fn add_export_fn(&mut self, interface: &str, name: &str, sig: super::component::FuncSig) {
        self.exports.push(ComponentExport {
            interface: interface.into(),
            name: name.into(),
            kind: ComponentItemKind::Function(sig),
        });
    }

    /// Add a resource export.
    pub fn add_export_resource(&mut self, interface: &str, name: &str, resource: ResourceType) {
        self.exports.push(ComponentExport {
            interface: interface.into(),
            name: name.into(),
            kind: ComponentItemKind::Resource(resource),
        });
    }

    /// Get all resource types (defined + imported + exported).
    pub fn all_resource_names(&self) -> Vec<&str> {
        let mut names: Vec<&str> = self.resources.iter().map(|r| r.name.as_str()).collect();
        for imp in &self.imports {
            if let ComponentItemKind::Resource(r) = &imp.kind {
                names.push(&r.name);
            }
        }
        for exp in &self.exports {
            if let ComponentItemKind::Resource(r) = &exp.kind {
                names.push(&r.name);
            }
        }
        names
    }
}

impl ResourceType {
    pub fn new(name: impl Into<String>) -> Self {
        ResourceType {
            name: name.into(),
            type_id: 0,
            constructor: None,
            methods: Vec::new(),
            destructor: None,
            borrowable: true,
        }
    }

    /// Add a method to this resource type.
    pub fn with_method(mut self, method: ResourceMethod) -> Self {
        self.methods.push(method);
        self
    }

    /// Set the constructor.
    pub fn with_constructor(mut self, ctor: ResourceMethod) -> Self {
        self.constructor = Some(ctor);
        self
    }

    /// Set the destructor.
    pub fn with_destructor(mut self, dtor: ResourceMethod) -> Self {
        self.destructor = Some(dtor);
        self
    }
}

impl ResourceMethod {
    pub fn new(name: impl Into<String>) -> Self {
        ResourceMethod {
            name: name.into(),
            is_static: false,
            borrows_self: true,
            params: Vec::new(),
            results: Vec::new(),
        }
    }

    pub fn static_method(name: impl Into<String>) -> Self {
        ResourceMethod {
            name: name.into(),
            is_static: true,
            borrows_self: false,
            params: Vec::new(),
            results: Vec::new(),
        }
    }

    pub fn with_params(mut self, params: Vec<super::component::ValType>) -> Self {
        self.params = params;
        self
    }

    pub fn with_results(mut self, results: Vec<super::component::ValType>) -> Self {
        self.results = results;
        self
    }

    pub fn consuming(mut self) -> Self {
        self.borrows_self = false;
        self
    }
}

/// Register a resource type in both the ComponentDescriptor and the VM's TypeRegistry.
/// Returns the type_id assigned by the registry.
pub fn register_resource_type(
    vm: &mut super::VM,
    resource: &mut ResourceType,
    parent_type_id: Option<usize>,
) -> usize {
    let mut typedef = super::TypeDef::new(&resource.name);
    if let Some(pid) = parent_type_id {
        typedef.parent = Some(pid);
    }
    let tid = vm.type_registry.register(typedef);
    resource.type_id = tid;
    tid
}

/// A resource table — tracks live resource handles for ownership/borrowing.
/// Each resource handle maps to its underlying data (stored as a Value).
#[derive(Debug, Default)]
pub struct ResourceTable {
    /// handle_id → (type_id, data, borrow_count)
    entries: HashMap<u32, ResourceEntry>,
    next_handle: u32,
}

#[derive(Debug, Clone)]
struct ResourceEntry {
    /// Type ID from TypeRegistry
    type_id: usize,
    /// The actual resource data (typically an Object with properties)
    data: super::Value,
    /// Number of active borrows (0 = can be dropped or moved)
    borrow_count: u32,
    /// Whether this resource has been dropped
    dropped: bool,
}

impl ResourceTable {
    pub fn new() -> Self {
        ResourceTable {
            entries: HashMap::new(),
            next_handle: 1,
        }
    }

    /// Create a new resource, returning its handle ID.
    pub fn create(&mut self, type_id: usize, data: super::Value) -> u32 {
        let handle = self.next_handle;
        self.next_handle += 1;
        self.entries.insert(handle, ResourceEntry {
            type_id,
            data,
            borrow_count: 0,
            dropped: false,
        });
        handle
    }

    /// Borrow a resource (read-only access). Returns the data if valid.
    pub fn borrow(&mut self, handle: u32) -> Option<super::Value> {
        let entry = self.entries.get_mut(&handle)?;
        if entry.dropped { return None; }
        entry.borrow_count += 1;
        Some(entry.data.clone())
    }

    /// Release a borrow.
    pub fn release_borrow(&mut self, handle: u32) {
        if let Some(entry) = self.entries.get_mut(&handle) {
            if entry.borrow_count > 0 {
                entry.borrow_count -= 1;
            }
        }
    }

    /// Drop a resource (destroy it). Fails if there are active borrows.
    pub fn drop_resource(&mut self, handle: u32) -> Result<super::Value, String> {
        let entry = self.entries.get_mut(&handle)
            .ok_or_else(|| format!("Invalid resource handle: {}", handle))?;
        if entry.dropped {
            return Err(format!("Resource {} already dropped", handle));
        }
        if entry.borrow_count > 0 {
            return Err(format!("Cannot drop resource {}: {} active borrows", handle, entry.borrow_count));
        }
        entry.dropped = true;
        Ok(entry.data.clone())
    }

    /// Get the type_id of a resource handle.
    pub fn type_id(&self, handle: u32) -> Option<usize> {
        self.entries.get(&handle).map(|e| e.type_id)
    }

    /// Check if a handle is valid and not dropped.
    pub fn is_valid(&self, handle: u32) -> bool {
        self.entries.get(&handle).map(|e| !e.dropped).unwrap_or(false)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resource_table_lifecycle() {
        let mut table = ResourceTable::new();

        // Create a resource
        let handle = table.create(1, crate::Value::String(std::rc::Arc::from("file_data")));
        assert!(table.is_valid(handle));
        assert_eq!(table.type_id(handle), Some(1));

        // Borrow it
        let data = table.borrow(handle).unwrap();
        assert!(matches!(data, crate::Value::String(_)));

        // Can't drop while borrowed
        assert!(table.drop_resource(handle).is_err());

        // Release borrow, then drop
        table.release_borrow(handle);
        let dropped = table.drop_resource(handle);
        assert!(dropped.is_ok());

        // Can't use after drop
        assert!(!table.is_valid(handle));
        assert!(table.borrow(handle).is_none());
    }

    #[test]
    fn test_component_descriptor() {
        use crate::component::{FuncSig, ValType};

        let mut comp = ComponentDescriptor::new("my-component");

        // Add a resource type
        let file_resource = ResourceType::new("FileHandle")
            .with_constructor(ResourceMethod::static_method("[constructor]")
                .with_params(vec![ValType::String])
                .with_results(vec![ValType::I32]))
            .with_method(ResourceMethod::new("read")
                .with_results(vec![ValType::String]))
            .with_method(ResourceMethod::new("write")
                .with_params(vec![ValType::String]))
            .with_destructor(ResourceMethod::new("[resource-drop]").consuming());

        comp.add_resource(file_resource);
        assert_eq!(comp.resources.len(), 1);
        assert_eq!(comp.resources[0].name, "FileHandle");
        assert_eq!(comp.resources[0].methods.len(), 2);
        assert!(comp.resources[0].constructor.is_some());
        assert!(comp.resources[0].destructor.is_some());

        // Add imports/exports
        comp.add_import_fn("wasi:io/streams", "read", FuncSig {
            name: "read".into(),
            params: vec![ValType::I32],
            results: vec![ValType::String],
        });

        comp.add_export_fn("my:api/math", "add", FuncSig {
            name: "add".into(),
            params: vec![ValType::I32, ValType::I32],
            results: vec![ValType::I32],
        });

        assert_eq!(comp.imports.len(), 1);
        assert_eq!(comp.exports.len(), 1);
        assert_eq!(comp.all_resource_names(), vec!["FileHandle"]);
    }
}
