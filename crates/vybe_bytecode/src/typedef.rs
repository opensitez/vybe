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
}

impl TypeDef {
    pub fn new(name: &str) -> Self {
        TypeDef {
            name: name.into(),
            parent: None,
            field_defs: Vec::new(),
            field_map: HashMap::new(),
            methods: HashMap::new(),
            constructor: None,
            destructor: None,
            is_resource: false,
            constants: HashMap::new(),
            exports: Vec::new(),
        }
    }

    pub fn with_parent(mut self, parent_id: usize) -> Self {
        self.parent = Some(parent_id);
        self
    }

    /// Add a named field. Returns the field index.
    pub fn add_field(&mut self, name: &str) -> usize {
        let idx = self.field_defs.len();
        let key = name.to_lowercase();
        self.field_defs.push(FieldDef { name: key.clone(), index: idx });
        self.field_map.insert(key, idx);
        idx
    }

    /// Get the index of a field by name, or None.
    pub fn field_index(&self, name: &str) -> Option<usize> {
        self.field_map.get(&name.to_lowercase()).copied()
    }

    /// Number of indexed fields (for Object.fields pre-allocation).
    pub fn field_count(&self) -> usize {
        self.field_defs.len()
    }

    pub fn method(mut self, name: &str, m: Method) -> Self {
        self.methods.insert(name.to_lowercase(), m);
        self
    }

    pub fn host_method(mut self, name: &str, host_fn_idx: usize) -> Self {
        self.methods.insert(name.to_lowercase(), Method::HostFn(host_fn_idx));
        self
    }

    pub fn resource(mut self) -> Self {
        self.is_resource = true;
        self
    }

    /// Backward compat: old code used `fields: Vec<String>` for field names.
    /// This preserves that interface.
    pub fn fields(&self) -> Vec<String> {
        self.field_defs.iter().map(|f| f.name.clone()).collect()
    }
}

/// Type registry — stores all type definitions.
/// Index 0 is reserved for "Object" (universal base type).
#[derive(Debug, Clone)]
pub struct TypeRegistry {
    pub types: Vec<TypeDef>,
    /// Name → type_id lookup
    name_map: HashMap<String, usize>,
}

impl TypeRegistry {
    pub fn new() -> Self {
        let mut reg = TypeRegistry {
            types: Vec::new(),
            name_map: HashMap::new(),
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

    /// Get type_id by name (case-insensitive).
    pub fn get_id(&self, name: &str) -> Option<usize> {
        self.name_map.get(&name.to_lowercase()).copied()
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
        let key = field_name.to_lowercase();
        let mut tid = type_id;
        loop {
            if let Some(typedef) = self.types.get(tid) {
                if let Some(&idx) = typedef.field_map.get(&key) {
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
        let key = method_name.to_lowercase();
        let mut tid = type_id;
        loop {
            if let Some(typedef) = self.types.get(tid) {
                if let Some(m) = typedef.methods.get(&key) {
                    return Some(m);
                }
                if let Some(parent) = typedef.parent {
                    tid = parent;
                } else {
                    // Try universal Object type (index 0)
                    if tid != 0 {
                        return self.types[0].methods.get(&key);
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
            typedef.methods.insert(method_name.to_lowercase(), m);
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
            typedef.constants.insert(name.to_lowercase(), value);
        }
    }

    /// Look up a constructor for a type.
    pub fn get_constructor(&self, type_id: usize) -> Option<&Method> {
        self.types.get(type_id)?.constructor.as_ref()
    }

    /// Look up a constant on a type (for enums).
    pub fn get_constant(&self, type_id: usize, name: &str) -> Option<i64> {
        let key = name.to_lowercase();
        self.types.get(type_id)?.constants.get(&key).copied()
    }

    /// Check if type_id is a subtype of target_id (walks parent chain).
    pub fn is_subtype(&self, type_id: usize, target_id: usize) -> bool {
        if type_id == target_id { return true; }
        let mut tid = type_id;
        loop {
            if let Some(typedef) = self.types.get(tid) {
                if let Some(parent) = typedef.parent {
                    if parent == target_id { return true; }
                    tid = parent;
                } else {
                    return false;
                }
            } else {
                return false;
            }
        }
    }

    /// Load type entries from a compiled chunk's type table.
    /// Called by VM.run() before execution. Registers user-defined types
    /// and adds their ChunkFn methods to the vtable.
    pub fn load_type_table(&mut self, types: &[super::chunk::TypeEntry]) {
        for entry in types {
            // Check if type already exists (e.g., host type registered earlier)
            let type_id = if let Some(existing) = self.get_id(&entry.name) {
                existing
            } else {
                let mut td = TypeDef::new(&entry.name);
                // Set parent
                if !entry.parent.is_empty() {
                    if let Some(pid) = self.get_id(&entry.parent) {
                        td.parent = Some(pid);
                    }
                }
                self.register(td)
            };

            // Add fields
            let typedef = &mut self.types[type_id];
            for field_name in &entry.fields {
                if typedef.field_index(field_name).is_none() {
                    typedef.add_field(field_name);
                }
            }

            // Add vtable methods (ChunkFn)
            for (method_name, chunk_idx) in &entry.methods {
                typedef.methods.insert(method_name.to_lowercase(), Method::ChunkFn(*chunk_idx));
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
        ResourceTable { entries: HashMap::new(), next_handle: 1 }
    }

    pub fn create(&mut self, type_id: usize, data: super::Value) -> u32 {
        let handle = self.next_handle;
        self.next_handle += 1;
        self.entries.insert(handle, ResourceEntry { type_id, data, borrow_count: 0, dropped: false });
        handle
    }

    pub fn borrow(&mut self, handle: u32) -> Option<super::Value> {
        let entry = self.entries.get_mut(&handle)?;
        if entry.dropped { return None; }
        entry.borrow_count += 1;
        Some(entry.data.clone())
    }

    pub fn release_borrow(&mut self, handle: u32) {
        if let Some(entry) = self.entries.get_mut(&handle) {
            if entry.borrow_count > 0 { entry.borrow_count -= 1; }
        }
    }

    pub fn drop_resource(&mut self, handle: u32) -> Result<super::Value, String> {
        let entry = self.entries.get_mut(&handle)
            .ok_or_else(|| format!("Invalid resource handle: {}", handle))?;
        if entry.dropped { return Err(format!("Resource {} already dropped", handle)); }
        if entry.borrow_count > 0 {
            return Err(format!("Cannot drop resource {}: {} active borrows", handle, entry.borrow_count));
        }
        entry.dropped = true;
        Ok(entry.data.clone())
    }

    pub fn type_id(&self, handle: u32) -> Option<usize> {
        self.entries.get(&handle).map(|e| e.type_id)
    }

    pub fn is_valid(&self, handle: u32) -> bool {
        self.entries.get(&handle).map(|e| !e.dropped).unwrap_or(false)
    }
}
