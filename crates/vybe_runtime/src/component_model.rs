//! Component Model resource and class patterns for WASM GC integration.
//!
//! This module extends the Component Model (component.rs) with richer typed
//! descriptors for both resources and classes. Resources follow the WASM
//! Component Model's `resource` pattern:
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

/// A class type definition in the Component Model.
/// Classes model user/framework objects with inheritance, fields,
/// properties, methods, and constructors.
///
/// This is the **wire format** — what crosses module boundaries in an
/// ESM-imported or Component-Model-linked class. The compile-time IR
/// (`NormalClass`) lives in `vybex::common::classes` and carries
/// additional metadata (spans, source names, special-method kinds,
/// event bindings) that is NOT part of the runtime description.
///
/// See `classnormalization.md` at the project root for the full
/// compile → runtime layering.
#[derive(Debug, Clone)]
pub struct ClassType {
    /// Class name (e.g. "Form", "Button", "StringBuilder")
    pub name: String,
    /// Optional parent class name. Resolved by name at register time;
    /// ESM dependency ordering guarantees the parent's `ClassType` is
    /// already in the `TypeRegistry` when this class is registered.
    pub parent: Option<String>,
    /// Interface / mixin names used for `instanceof` / `isinstance` /
    /// `is` / `kind_of?`. Mixin + trait methods are **flattened into
    /// `methods` at walker time** — this list is only for identity
    /// checks, not dispatch.
    pub interfaces: Vec<String>,
    /// Plain instance fields. Field index = position in this Vec;
    /// ordering is part of the wire contract so cross-module code
    /// observes the same layout.
    pub fields: Vec<String>,
    /// Properties, potentially backed by host getter/setter calls.
    pub properties: Vec<PropertyDef>,
    /// Instance or static methods, keyed on canonical (language-neutral)
    /// method names.
    ///
    /// **Canonical means canonical — there is no alias fallback.** A
    /// `method_aliases` table sat beside this, mapping `"__str__"`/`"to_s"` to
    /// `"tostring"`, and it had ZERO readers anywhere in the workspace:
    /// declared here, populated nowhere, consulted by nothing.
    /// `classnormalization.md` says the walker resolves the alias when it
    /// produces `NormalClass`, so by the time a name reaches the runtime it is
    /// already canonical and the table could only ever have been dead. Same
    /// failure family as flexclassplan §1b's `cross_language_aliases`, which
    /// was deleted for the same reason.
    pub methods: Vec<MethodDef>,
    /// Constructor definitions, in registration order.
    ///
    /// ⛔ A LIST, BECAUSE A TYPE HAS OVERLOADED CONSTRUCTORS. This was a single
    /// `Option` and `with_constructor` OVERWROTE, so every type registering more
    /// than one kept only the last — invisible to dispatch, which asks the tree
    /// for one backing, and fatal to reflection, which must match
    /// `GetConstructor(Type[])` against the real set.
    ///
    /// [`ClassType::constructor`] answers the LAST, which is exactly what the
    /// overwriting field held, so dispatch is unchanged by the widening.
    pub constructors: Vec<ConstructorDef>,
    /// Optional destructor/finalizer-like method.
    pub destructor: Option<MethodDef>,
}

/// A property on a class type.
#[derive(Debug, Clone)]
pub struct PropertyDef {
    pub name: String,
    pub setter: Option<HostTarget>,
    pub getter: Option<HostTarget>,
}

/// A method on a class type.
#[derive(Debug, Clone)]
pub struct MethodDef {
    pub name: String,
    pub is_static: bool,
    /// User-visible arity. For instance methods this excludes the implicit `this`.
    pub arity: u8,
    pub body: MethodBody,
    /// Declared parameter types, in Component Model `ValType`s.
    ///
    /// An arity says how MANY, never WHAT — and a descriptor that states only a
    /// count is not a typed interface member. Overload selection by type
    /// (`GetMethod(name, Type[])`) and reflection's `ParameterType` both need
    /// the types, so a platform that has them declares them here rather than in
    /// a table beside the descriptor: the seeder reads descriptors as DATA
    /// across a crate boundary and cannot call back into the platform.
    ///
    /// `None` means undeclared — the arity is all this leaf carries, which is
    /// what every registration starts as.
    pub params: Option<Vec<super::component::ValType>>,
    /// Declared result type, when the platform states one.
    pub result: Option<super::component::ValType>,
}

/// The body for a class method.
#[derive(Debug, Clone)]
pub enum MethodBody {
    /// Method implemented by an emitted chunk index.
    UserChunk(usize),
    /// Method forwards to a host function.
    HostCall(HostTarget),
    /// Method lowers through a shared compiler-side emitter.
    Common(String),
}

/// Constructor definition for a class type.
#[derive(Debug, Clone)]
pub struct ConstructorDef {
    pub arity: u8,
    /// Optional backing that materializes the runtime object.
    pub backing: Option<ConstructorTarget>,
    /// Declared parameter types. See [`MethodDef::params`].
    pub params: Option<Vec<super::component::ValType>>,
}

/// A constructor backing can either be a host import or a canonical
/// compiler-side common emit.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ConstructorTarget {
    Host(HostTarget),
    Common(String),
}

/// A host target resolved by the linker against registered host exports.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HostTarget {
    pub module: String,
    pub name: String,
}

/// A component import/export kind — extends the basic Component with resource support.
#[derive(Debug, Clone)]
pub enum ComponentItemKind {
    /// A function import/export
    Function(super::component::FuncSig),
    /// A class type import/export
    Class(ClassType),
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
    /// Class types defined by this component.
    pub classes: Vec<ClassType>,
    /// Resource types defined by this component
    pub resources: Vec<ResourceType>,
}

impl ComponentDescriptor {
    pub fn new(name: impl Into<String>) -> Self {
        ComponentDescriptor {
            name: name.into(),
            imports: Vec::new(),
            exports: Vec::new(),
            classes: Vec::new(),
            resources: Vec::new(),
        }
    }

    /// Add a class type definition to this component.
    pub fn add_class(&mut self, class: ClassType) {
        self.classes.push(class);
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

    /// Add a class import.
    pub fn add_import_class(&mut self, interface: &str, name: &str, class: ClassType) {
        self.imports.push(ComponentImport {
            interface: interface.into(),
            name: name.into(),
            kind: ComponentItemKind::Class(class),
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

    /// Add a class export.
    pub fn add_export_class(&mut self, interface: &str, name: &str, class: ClassType) {
        self.exports.push(ComponentExport {
            interface: interface.into(),
            name: name.into(),
            kind: ComponentItemKind::Class(class),
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

    /// Get all class types (defined + imported + exported).
    pub fn all_class_names(&self) -> Vec<&str> {
        let mut names: Vec<&str> = self.classes.iter().map(|c| c.name.as_str()).collect();
        for imp in &self.imports {
            if let ComponentItemKind::Class(class) = &imp.kind {
                names.push(&class.name);
            }
        }
        for exp in &self.exports {
            if let ComponentItemKind::Class(class) = &exp.kind {
                names.push(&class.name);
            }
        }
        names
    }
}

impl ClassType {
    pub fn new(name: impl Into<String>) -> Self {
        ClassType {
            name: name.into(),
            parent: None,
            interfaces: Vec::new(),
            fields: Vec::new(),
            properties: Vec::new(),
            methods: Vec::new(),
            constructors: Vec::new(),
            destructor: None,
        }
    }

    /// The constructor dispatch uses — the LAST registered, which is what the
    /// overwriting `Option` field held.
    pub fn constructor(&self) -> Option<&ConstructorDef> {
        self.constructors.last()
    }

    pub fn with_parent(mut self, parent: impl Into<String>) -> Self {
        self.parent = Some(parent.into());
        self
    }

    /// Declare an implemented interface / mixin / trait name used by
    /// `instanceof`-family checks. Method dispatch does NOT walk this
    /// list — mixin methods are flattened into `methods` at walker time.
    pub fn with_interface(mut self, interface: impl Into<String>) -> Self {
        self.interfaces.push(interface.into());
        self
    }

    pub fn with_field(mut self, field: impl Into<String>) -> Self {
        self.fields.push(field.into());
        self
    }

    pub fn with_property(mut self, property: PropertyDef) -> Self {
        self.properties.push(property);
        self
    }

    pub fn with_method(mut self, method: MethodDef) -> Self {
        self.methods.push(method);
        self
    }

    pub fn with_constructor(mut self, constructor: ConstructorDef) -> Self {
        self.constructors.push(constructor);
        self
    }

    pub fn with_destructor(mut self, destructor: MethodDef) -> Self {
        self.destructor = Some(destructor);
        self
    }
}

impl PropertyDef {
    pub fn new(name: impl Into<String>) -> Self {
        PropertyDef {
            name: name.into(),
            setter: None,
            getter: None,
        }
    }

    pub fn with_setter(mut self, setter: HostTarget) -> Self {
        self.setter = Some(setter);
        self
    }

    pub fn with_getter(mut self, getter: HostTarget) -> Self {
        self.getter = Some(getter);
        self
    }
}

impl MethodDef {
    pub fn new(name: impl Into<String>, arity: u8, body: MethodBody) -> Self {
        MethodDef {
            name: name.into(),
            is_static: false,
            arity,
            body,
            params: None,
            result: None,
        }
    }

    pub fn static_method(name: impl Into<String>, arity: u8, body: MethodBody) -> Self {
        MethodDef {
            name: name.into(),
            is_static: true,
            arity,
            body,
            params: None,
            result: None,
        }
    }

    /// Declare this method's parameter types. The arity is taken from them, so
    /// the count and the types cannot disagree.
    pub fn with_params(mut self, params: Vec<super::component::ValType>) -> Self {
        self.arity = params.len() as u8;
        self.params = Some(params);
        self
    }

    /// Declare this method's result type.
    pub fn with_result(mut self, result: super::component::ValType) -> Self {
        self.result = Some(result);
        self
    }
}

impl ConstructorDef {
    /// Declare this constructor's parameter types. See [`MethodDef::with_params`].
    pub fn with_params(mut self, params: Vec<super::component::ValType>) -> Self {
        self.arity = params.len() as u8;
        self.params = Some(params);
        self
    }

    pub fn new(arity: u8) -> Self {
        ConstructorDef {
            arity,
            backing: None,
            params: None,
        }
    }

    pub fn with_backing(mut self, backing: HostTarget) -> Self {
        self.backing = Some(ConstructorTarget::Host(backing));
        self
    }

    pub fn with_common_backing(mut self, emit: impl Into<String>) -> Self {
        self.backing = Some(ConstructorTarget::Common(emit.into()));
        self
    }
}

impl HostTarget {
    pub fn new(module: impl Into<String>, name: impl Into<String>) -> Self {
        HostTarget {
            module: module.into(),
            name: name.into(),
        }
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

    /// Borrow a resource (read-only access). Returns the data if valid.
    pub fn borrow(&mut self, handle: u32) -> Option<super::Value> {
        let entry = self.entries.get_mut(&handle)?;
        if entry.dropped {
            return None;
        }
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

    /// Get the type_id of a resource handle.
    pub fn type_id(&self, handle: u32) -> Option<usize> {
        self.entries.get(&handle).map(|e| e.type_id)
    }

    /// Check if a handle is valid and not dropped.
    pub fn is_valid(&self, handle: u32) -> bool {
        self.entries
            .get(&handle)
            .map(|e| !e.dropped)
            .unwrap_or(false)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resource_table_lifecycle() {
        let mut table = ResourceTable::new();

        // Create a resource
        use std::sync::Arc;
        let handle = table.create(1, crate::Value::String(Arc::from("file_data")));
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
            .with_constructor(
                ResourceMethod::static_method("[constructor]")
                    .with_params(vec![ValType::String])
                    .with_results(vec![ValType::I32]),
            )
            .with_method(ResourceMethod::new("read").with_results(vec![ValType::String]))
            .with_method(ResourceMethod::new("write").with_params(vec![ValType::String]))
            .with_destructor(ResourceMethod::new("[resource-drop]").consuming());

        comp.add_resource(file_resource);
        assert_eq!(comp.resources.len(), 1);
        assert_eq!(comp.resources[0].name, "FileHandle");
        assert_eq!(comp.resources[0].methods.len(), 2);
        assert!(comp.resources[0].constructor.is_some());
        assert!(comp.resources[0].destructor.is_some());

        let button_class = ClassType::new("Button")
            .with_parent("Control")
            .with_field("Text")
            // A made-up host module. This exercises `ClassType`'s SHAPE — a
            // property with a setter, a method with a host call, a backed
            // constructor — and nothing here resolves against a real registry.
            // ⚠ It must not name a real module: a fixture pointing at one
            // reads like a surviving dependency, to a reader and to grep.
            .with_property(
                PropertyDef::new("Enabled")
                    .with_setter(HostTarget::new("test:widgets", "setProperty")),
            )
            .with_method(MethodDef::new(
                "PerformClick",
                0,
                MethodBody::HostCall(HostTarget::new("test:widgets", "performClick")),
            ))
            .with_constructor(
                ConstructorDef::new(0).with_backing(HostTarget::new("test:widgets", "newButton")),
            );

        comp.add_class(button_class.clone());
        assert_eq!(comp.classes.len(), 1);
        assert_eq!(comp.classes[0].name, "Button");
        assert_eq!(comp.classes[0].parent.as_deref(), Some("Control"));
        assert_eq!(comp.classes[0].properties.len(), 1);
        assert_eq!(comp.classes[0].methods.len(), 1);

        // Add imports/exports
        comp.add_import_fn(
            "wasi:io/streams",
            "read",
            FuncSig {
                name: "read".into(),
                params: vec![ValType::I32],
                results: vec![ValType::String],
            },
        );

        comp.add_export_fn(
            "my:api/math",
            "add",
            FuncSig {
                name: "add".into(),
                params: vec![ValType::I32, ValType::I32],
                results: vec![ValType::I32],
            },
        );

        comp.add_import_class(
            "my:ui/widgets",
            "Button",
            button_class.clone(),
        );
        comp.add_export_class("my:ui/widgets", "Button", button_class);

        assert_eq!(comp.imports.len(), 2);
        assert_eq!(comp.exports.len(), 2);
        assert_eq!(comp.all_resource_names(), vec!["FileHandle"]);
        assert_eq!(comp.all_class_names(), vec!["Button", "Button", "Button"]);
    }
}

// ── Descriptor lookup targets ───────────────────────────────────────────
//
// What a component-descriptor lookup resolves to. These live here, beside
// `ConstructorTarget`, because the COMPILER receives them across a registry
// function pointer — a platform-local type cannot cross that boundary, and
// that is precisely why `vybe_compiler` still had a Cargo dependency on
// `vybe_platform_dotnet`.

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StaticPropertyTarget {
    Host { module: String, func: String },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum InstanceMethodTarget {
    Host {
        module: String,
        func: String,
        arity: u8,
    },
    Common {
        emit: String,
        arity: u8,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum InstancePropertyTarget {
    /// A host-backed property accessor. `key` is `Some(PascalName)` when the
    /// target is a *generic* property host fn
    /// (`getProperty(this, "Text")` / `setProperty(this, "Text", value)`) — the
    /// compiler pushes the key as an argument. `None` for dedicated per-property host fns
    /// (`Environment.NewLine` → `node:os.EOL(this)`).
    Host {
        module: String,
        func: String,
        key: Option<String>,
    },
    Common {
        emit: String,
    },
}
