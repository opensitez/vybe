//! WASM GC-style type definitions.
//!
//! Each TypeDef has:
//! - A name (for debugging / TypeOf)
//! - A parent type (for inheritance)
//! - A method table (vtable): method_name → host_fn_index or chunk_index
//! - Field names (for struct layout)
//!
//! Objects carry a `type_id` (index into VM's type_defs).
//! Method dispatch: check vtable on the object's type → parent type → universal.

use std::collections::HashMap;

/// A method entry in a type's vtable.
#[derive(Debug, Clone)]
pub enum Method {
    /// Host function by index in VM's host_fns table.
    HostFn(usize),
    /// Bytecode function chunk index.
    ChunkFn(usize),
}

/// A type definition — like a WASM GC struct type.
#[derive(Debug, Clone)]
pub struct TypeDef {
    /// Type name (e.g. "List", "Dictionary", "String", "DateTime")
    pub name: String,
    /// Parent type index (for inheritance). None = no parent.
    pub parent: Option<usize>,
    /// Method table: lowercase method name → Method
    pub methods: HashMap<String, Method>,
    /// Field names (optional, for documentation)
    pub fields: Vec<String>,
    /// Constructor method (called by `New TypeName()`)
    pub constructor: Option<Method>,
    /// Enum constants: name → value (for compile-time resolution)
    pub constants: HashMap<String, i64>,
}

impl TypeDef {
    pub fn new(name: &str) -> Self {
        TypeDef {
            name: name.into(),
            parent: None,
            methods: HashMap::new(),
            fields: Vec::new(),
            constructor: None,
            constants: HashMap::new(),
        }
    }

    pub fn with_parent(mut self, parent_id: usize) -> Self {
        self.parent = Some(parent_id);
        self
    }

    pub fn method(mut self, name: &str, m: Method) -> Self {
        self.methods.insert(name.to_lowercase(), m);
        self
    }

    pub fn host_method(mut self, name: &str, host_fn_idx: usize) -> Self {
        self.methods.insert(name.to_lowercase(), Method::HostFn(host_fn_idx));
        self
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

    /// Look up a method on a type, walking up the inheritance chain.
    pub fn resolve_method(&self, type_id: usize, method_name: &str) -> Option<&Method> {
        let key = method_name.to_lowercase();
        let mut tid = type_id;
        loop {
            if let Some(typedef) = self.types.get(tid) {
                if let Some(m) = typedef.methods.get(&key) {
                    return Some(m);
                }
                // Walk to parent
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

    /// Add a host method by looking up the host registry.
    pub fn add_host_method(&mut self, type_id: usize, method_name: &str, host_fn_idx: usize) {
        self.add_method(type_id, method_name, Method::HostFn(host_fn_idx));
    }

    /// Set constructor for a type.
    pub fn set_constructor(&mut self, type_id: usize, m: Method) {
        if let Some(typedef) = self.types.get_mut(type_id) {
            typedef.constructor = Some(m);
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
}
