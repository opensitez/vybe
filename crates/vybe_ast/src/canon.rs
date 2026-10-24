//! The canon section — a component's canonical definitions.
//!
//! `proposals/component-model/design/mvp/Binary.md` §"Canonical Definitions" is
//! not a list of opcodes; it is a section of **typed definition records**, and
//! almost every row carries immediates:
//!
//! ```text
//! (canon lift  $f opts $ft)        (canon lower $f opts)
//! (canon stream.read  $t opts)     (canon task.return (result $t)? opts)
//! (canon context.get  $v $i)       (canon thread.yield cancellable?)
//! (canon thread.new-indirect $ft $tbl)
//! ```
//!
//! # Why this is stated here as well as in the runtime
//!
//! `vybe_runtime::canon_def::CanonDef` is the same table as the runtime reads
//! it. This is the same table as the SOURCE states it, and the two are
//! deliberately different types:
//!
//! * an index here is whatever the component's own index spaces say, resolved
//!   by the front end from `$id` to a number;
//! * a `string-encoding` here is the token the source wrote, not the runtime's
//!   `StringEncoding`;
//! * a `context.get` valtype here is the primitive's SPELLING, because the
//!   runtime's `ValType` is a recursive type the AST has no business restating.
//!
//! That split is the one `Chunk::globals` already uses: the front end states
//! what the source said, and the compiler converts at emission. `vybe_runtime`
//! depends on `vybe_ast`, so the runtime's type could not be used here anyway
//! without dragging `component::ValType` and its transitive surface down a
//! layer.
//!
//! # Why it is not a `Statement`
//!
//! A canon definition has no execution position. It is not ordered against
//! code, it cannot be branched over, and its index space is module-level — the
//! same category as the global index space, which is likewise not an AST node.

use std::sync::Arc;

/// What a `canon lift` / `canon lower` wraps.
///
/// The sort is carried, never inferred. `canon lift $callee` takes a **core**
/// funcidx; `canon lower $callee` takes a **component** funcidx. A bare `u32`
/// makes those indistinguishable, and a mis-wired definition then lifts with
/// core types and looks entirely correct.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CanonCallee {
    /// `core:funcidx` as a BARE NUMBER, written positionally in the source.
    ///
    /// The runtime reads this as a chunk index. That is a different space from
    /// the component's core function index space, which is why a NAMED callee
    /// cannot use this variant — see [`Self::CoreExport`].
    Core(u32),
    /// `funcidx` in the component's function index space — `canon lower`.
    Component(u32),
    /// A core function named by the core instance export it came from.
    ///
    /// This is what `(canon lift (core func $r) …)` produces once `$r` has been
    /// resolved through the component's core function index space. It carries
    /// NAMES, not a number, because the number the runtime wants is a chunk
    /// index that only the compiler assigns — and writing the front end's index
    /// there would call whichever chunk happened to sit at it.
    ///
    /// The compiler resolves it against the type table's vtable
    /// (`TypeEntry.methods`, which is exactly `(class, method) -> chunk`) at
    /// install time, once every chunk exists.
    CoreExport { class: String, method: String },
}

/// A core function named in a canonical option. Source indices and compiled
/// chunk indices are different spaces, so named exports resolve after emission.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CanonCoreFuncRef {
    Index(u32),
    CoreExport { class: String, method: String },
}

/// `opts ::= opt*:vec(<canonopt>)`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CanonOptions {
    /// The token as written: `"utf8"`, `"utf16"`, `"latin1+utf16"`.
    ///
    /// `None` is NOT `utf8`. The spec's default is utf8, but "not stated" and
    /// "stated as utf8" are different facts about the source, and only the
    /// former may be overridden by a surrounding declaration.
    pub string_encoding: Option<String>,
    /// `(memory m)` — core memidx.
    pub memory: Option<u32>,
    /// `(realloc f)` — core funcidx.
    pub realloc: Option<CanonCoreFuncRef>,
    /// `(post-return f)` — core funcidx, `canon lift` only.
    pub post_return: Option<CanonCoreFuncRef>,
    /// `(callback f)` — core funcidx, 🔀 async lift only.
    pub callback: Option<CanonCoreFuncRef>,
    /// `async` 🔀.
    pub is_async: bool,
}

/// One row of the canon section, as the source states it.
///
/// Every immediate any row can carry, in one record. Rows that take none leave
/// the fields at their defaults. Nothing here is defaulted on the way IN: a row
/// that omits an immediate leaves `None`, so the consumer can tell "not
/// declared" from "declared as zero". A canonical ABI that is quietly wrong
/// about layout corrupts a peer's memory with nothing to detect it.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CanonDecl {
    /// The row's spec spelling — `"lift"`, `"lower"`, `"thread.suspend"`,
    /// `"stream.read"`. Carried as written so the error a bad row produces can
    /// quote the source rather than a rule name.
    pub builtin: String,
    /// The `(core func <id>?)` this row binds, if it names one. This is what
    /// makes a canon definition reachable: a core module inside the component
    /// imports the binder, and the call lands on the built-in.
    pub binder: Option<String>,
    /// `$callee` — `lift` / `lower` only.
    pub callee: Option<CanonCallee>,
    /// `t:<typeidx>` (stream/future element) or `rt:<typeidx>` (resource).
    pub ty: Option<u32>,
    /// `ft` as a FUNCTION type, for `canon lift`. Separate from `ty` because
    /// they index different spaces: a value type and a function type.
    pub functype: Option<u32>,
    /// `tbl:<core:tableidx>` — `thread.new-indirect`, `thread.spawn-indirect`.
    pub table: Option<u32>,
    /// Resolved core worker signature: exactly one i32/i64 parameter, no results.
    /// Separate from the component type space indexed by `ty` on other rows.
    pub worker_param: Option<ComponentCoreValType>,
    /// `rs:<resultlist>` — `canon task.return`.
    pub results: Vec<u32>,
    /// `(v, i)` — `context.get` / `context.set`. The valtype is the primitive's
    /// SPELLING (`"i32"`, `"u32"`); the compiler resolves it and refuses a
    /// spelling it cannot name. Two immediates, which is precisely what a
    /// single `Option<u32>` channel could never carry.
    pub context: Option<(String, u32)>,
    pub opts: CanonOptions,
    /// `cancellable?` — `thread.{suspend,yield,…}`, `waitable-set.{wait,poll}`.
    pub cancellable: bool,
    /// `async?` — `subtask.cancel`, `{stream,future}.cancel-*`.
    pub is_async: bool,
    /// `shared?` 🧵② — the spawn rows.
    pub shared: bool,
}

impl CanonDecl {
    /// A row that carries nothing but its name and binder.
    pub fn new(builtin: impl Into<String>, binder: Option<String>) -> Self {
        CanonDecl {
            builtin: builtin.into(),
            binder,
            ..Default::default()
        }
    }
}

/// A value type, as the SOURCE spells it — `Explainer.md` §8 `defvaltype`.
///
/// Deliberately NOT `vybe_runtime::component::ValType`. The runtime's type is
/// what the ABI can represent; this is what the source is allowed to say, and
/// the two are not the same set. The spec has eight integer widths, `char`,
/// `f32`, `tuple`, `flags` and `enum`; the runtime has one `I32`, one `I64`
/// and no `char`. Recording the SPELLING keeps that difference visible so the
/// compiler can refuse a width it cannot carry, instead of mapping `u8` onto
/// `I32` and truncating nothing where the spec requires truncation.
///
/// Same split as `CanonDecl` -> `CanonDef`: the front end states what the
/// source said, the compiler converts at emission and refuses what it cannot
/// name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ValSpec {
    /// A primitive, as written: `"u32"`, `"s8"`, `"char"`, `"string"`, `"bool"`,
    /// `"f32"`, `"error-context"` …
    Prim(String),
    List(Box<ValSpec>),
    /// 🔧 fixed-length list — `(list <valtype> <u32>)`.
    ListFixed(Box<ValSpec>, u32),
    Record(Vec<(String, ValSpec)>),
    Variant(Vec<(String, Option<ValSpec>)>),
    Tuple(Vec<ValSpec>),
    Flags(Vec<String>),
    Enum(Vec<String>),
    Option(Box<ValSpec>),
    /// `result<T, E>` — each side independently optional.
    Result(Option<Box<ValSpec>>, Option<Box<ValSpec>>),
    /// 🗺️ `(map <keytype> <valtype>)`.
    Map(Box<ValSpec>, Box<ValSpec>),
    /// 🔀 `(stream <valtype>?)` — the element type is optional.
    Stream(Option<Box<ValSpec>>),
    /// 🔀 `(future <valtype>?)`.
    Future(Option<Box<ValSpec>>),
    /// `(own <typeidx>)` / `(borrow <typeidx>)` — a resource handle, named by
    /// INDEX in the source. The runtime names resources by string, so the
    /// compiler has to resolve this against the type space.
    Own(u32),
    Borrow(u32),
    /// A type named by index — `(type $t)` used as a valtype.
    Ref(u32),
}

/// One entry in a component's TYPE index space, as the source states it.
///
/// The index is positional and every `(type …)` advances it, named or not, so
/// a form this front end does not decompose is still recorded as `Opaque`
/// rather than skipped. Skipping one would silently renumber every later index
/// — the same reason `walk_comp_type` counts unnamed declarations.
#[derive(Debug, Clone, PartialEq)]
pub enum TypeDecl {
    /// `(type $t (func (param "a" u32) (result u32)))`.
    ///
    /// The canonical ABI is single-result (`MAX_FLAT_RESULTS` is 1), which is
    /// why `result` is an `Option` and not a `Vec`.
    Func {
        params: Vec<(String, ValSpec)>,
        result: Option<ValSpec>,
        /// Component function effect, independent of the canonical ABI option.
        is_async: bool,
    },
    /// `(type $t u32)`, `(type $t (list string))` — a VALUE type.
    Value(ValSpec),
    /// A `resourcetype`, `componenttype` or `instancetype`. It occupies an
    /// index; nothing yet reads its shape beyond a resource's NAME. The string
    /// is the form's keyword so a refusal can say which one it met.
    Opaque(String),
    /// `(type $t (resource (rep i32) (dtor ...)))` — a resource's bound name,
    /// core representation width, and optional destructor.
    ///
    /// Separate from [`Self::Opaque`] because `own`/`borrow` need the NAME,
    /// not the keyword. `component::ValType::Own` is `Own(String)` — a handle
    /// binds to whatever is registered under that name — while the source
    /// names the resource by INDEX. This is the only record of which name that
    /// index belongs to, so folding it back into `Opaque("resource")` (which
    /// is what it used to be) makes every `(own $r)` unresolvable.
    ///
    /// An UNNAMED resource type keeps `None`: it occupies its index, and a
    /// handle to it refuses rather than being given an invented name.
    Resource {
        name: Option<String>,
        rep: ResourceRep,
        dtor: Option<CanonCallee>,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ResourceRep {
    I32,
    I64,
}

/// The four outeraliassort namespaces, not the larger set of export sorts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComponentOuterAliasKind {
    CoreModule,
    CoreType,
    Type,
    Component,
}

/// `count == 0` selects the current scope; `index` belongs to `kind`'s namespace.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ComponentOuterAlias {
    pub kind: ComponentOuterAliasKind,
    pub count: u32,
    pub index: u32,
}

impl ComponentBinaryDef {
    pub fn normalized_outer_alias(&self) -> Option<ComponentOuterAlias> {
        match self {
            Self::AliasOuter(alias) => Some(*alias),
            Self::CoreAliasOuterType { count, index } => Some(ComponentOuterAlias {
                kind: ComponentOuterAliasKind::CoreType,
                count: *count,
                index: *index,
            }),
            _ => None,
        }
    }
}

/// Ordered definitions of a component binary. Core modules retain their
/// validated binary form; their index space is separate from core instances.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ComponentBinary {
    pub definitions: Vec<ComponentBinaryDef>,
}

#[path = "component_drop.rs"]
mod component_drop;

/// Ordered name metadata; duplicate kinds remain available for validation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ComponentNameAttribute {
    Implements(String),
    VersionSuffix(String),
    ExternalId(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComponentExternalName {
    pub name: String,
    pub attributes: Vec<ComponentNameAttribute>,
}

impl From<String> for ComponentExternalName {
    fn from(name: String) -> Self { Self { name, attributes: Vec::new() } }
}
impl From<&str> for ComponentExternalName {
    fn from(name: &str) -> Self { name.to_owned().into() }
}

#[derive(Debug, Clone, Copy)]
pub struct ComponentExternalNameRef<'a> {
    pub name: &'a str,
    pub attributes: &'a [ComponentNameAttribute],
}
impl ComponentExternalName {
    pub fn as_ref(&self) -> ComponentExternalNameRef<'_> {
        ComponentExternalNameRef { name: &self.name, attributes: &self.attributes }
    }
}
impl<'a> From<&'a str> for ComponentExternalNameRef<'a> {
    fn from(name: &'a str) -> Self { Self { name, attributes: &[] } }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComponentNamedItemRef {
    pub name: ComponentExternalName,
    pub kind: ComponentItemKind,
    pub index: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ComponentBinaryDef {
    CoreType(ComponentCoreType),
    AliasOuter(ComponentOuterAlias),
    /// Compatibility constructor; use `normalized_outer_alias` when consuming.
    CoreAliasOuterType { count: u32, index: u32 },
    InterfaceType {
        is_component: bool,
        declarations: Vec<ComponentInterfaceDecl>,
    },
    /// Compatibility constructor; codecs and validation normalize to `CoreType`.
    CoreFuncType {
        params: Vec<ComponentCoreValType>,
        results: Vec<ComponentCoreValType>,
    },
    Custom {
        name: String,
        data: Vec<u8>,
    },
    ValueType(ValSpec),
    ValueLiteral { ty: ValSpec, payload: Vec<u8> },
    Start { function: u32, args: Vec<u32>, results: u32 },
    FuncType {
        params: Vec<(String, ValSpec)>,
        result: Option<ValSpec>,
        is_async: bool,
    },
    ResourceType {
        rep: ResourceRep,
        dtor: Option<u32>,
    },
    CanonLift {
        core_func: u32,
        type_index: u32,
        options: Vec<ComponentBinaryCanonOption>,
    },
    CanonLower {
        func: u32,
        options: Vec<ComponentBinaryCanonOption>,
    },
    CanonTaskReturn {
        result: Option<ValSpec>,
        options: Vec<ComponentBinaryCanonOption>,
    },
    CanonAsync {
        kind: ComponentBinaryAsyncKind,
        operation: ComponentBinaryAsyncOperation,
        type_index: u32,
        options: Vec<ComponentBinaryCanonOption>,
    },
    CanonThreadIndex,
    CanonResourceNew { type_index: u32 },
    CanonResourceDrop { type_index: u32 },
    CanonResourceRep { type_index: u32 },
    CanonErrorContextNew { options: Vec<ComponentBinaryCanonOption> },
    CanonErrorContextDebugMessage { options: Vec<ComponentBinaryCanonOption> },
    CanonErrorContextDrop,
    CanonWaitableSetNew,
    CanonWaitableSetWait { memory: u32 },
    CanonWaitableSetPoll { memory: u32 },
    CanonWaitableSetDrop,
    CanonWaitableJoin,
    CanonContext { set: bool, ty: ComponentCoreValueType, slot: u32 },
    CanonTaskCancel,
    CanonSubtaskCancel { is_async: bool },
    CanonSubtaskDrop,
    CanonBackpressureInc,
    CanonBackpressureDec,
    CanonThreadResumeLater,
    CanonThreadSuspend,
    CanonThreadYield,
    CanonThreadSuspendThenResume,
    CanonThreadYieldThenResume,
    CanonThreadSuspendThenPromote,
    CanonThreadYieldThenPromote,
    CanonThreadNewIndirect {
        type_index: u32,
        table_index: u32,
    },
    CanonThreadSpawnRef {
        shared: bool,
        type_index: u32,
    },
    CanonThreadSpawnIndirect {
        shared: bool,
        type_index: u32,
        table_index: u32,
    },
    CanonThreadAvailableParallelism {
        shared: bool,
    },
    ImportFunc {
        name: String,
        type_index: u32,
    },
    Import {
        name: String,
        ty: ComponentInterfaceRef,
    },
    NamedImport {
        name: ComponentExternalName,
        ty: ComponentInterfaceRef,
    },
    NamedExport {
        name: ComponentExternalName,
        kind: ComponentItemKind,
        index: u32,
        ty: Option<ComponentInterfaceRef>,
    },
    ExportFunc {
        name: String,
        index: u32,
        /// Optional function-type ascription in the component type space.
        type_index: Option<u32>,
    },
    NestedComponent(Arc<ComponentBinary>),
    ComponentInstance {
        component: u32,
        args: Vec<ComponentItemRef>,
    },
    InstanceFromExports(Vec<ComponentItemRef>),
    NamedInstanceFromExports(Vec<ComponentNamedItemRef>),
    AliasInstanceExport {
        instance: u32,
        name: String,
        kind: ComponentItemKind,
    },
    ExportType {
        name: String,
        index: u32,
    },
    ExportCoreModule {
        name: String,
        index: u32,
    },
    ExportInstance {
        name: String,
        index: u32,
    },
    CoreModule(Vec<u8>),
    CoreInstance {
        module: u32,
        /// Import module name and the supplying core instance index.
        args: Vec<(String, u32)>,
    },
    CoreInstanceFromExports(Vec<CoreInstanceExport>),
    CoreAliasInstanceExport {
        instance: u32,
        name: String,
        kind: CoreExportKind,
    },
}

/// Ordered declarations inside a component or instance type. Type indices are
/// local to this body; outer aliases explicitly cross its lexical boundary.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ComponentInterfaceDecl {
    CoreType(ComponentCoreType),
    AliasOuterCoreType { count: u32, index: u32 },
    Type(Box<ComponentBinaryDef>),
    AliasOuterType { count: u32, index: u32 },
    AliasInstanceExport { instance: u32, name: String, kind: ComponentItemKind },
    Import { name: String, ty: ComponentInterfaceRef },
    Export { name: String, ty: ComponentInterfaceRef },
    NamedImport { name: ComponentExternalName, ty: ComponentInterfaceRef },
    NamedExport { name: ComponentExternalName, ty: ComponentInterfaceRef },
}

impl ComponentInterfaceDecl {
    /// String-based constructors normalize to an empty attribute list.
    pub fn external(&self) -> Option<(bool, ComponentExternalNameRef<'_>, &ComponentInterfaceRef)> {
        match self {
            Self::Import { name, ty } => Some((true, name.as_str().into(), ty)),
            Self::Export { name, ty } => Some((false, name.as_str().into(), ty)),
            Self::NamedImport { name, ty } => Some((true, name.as_ref(), ty)),
            Self::NamedExport { name, ty } => Some((false, name.as_ref(), ty)),
            _ => None,
        }
    }
}

impl ComponentBinaryDef {
    pub fn external_name(&self) -> Option<(ComponentExternalNameRef<'_>, bool)> {
        match self {
            Self::ImportFunc { name, .. } | Self::ExportFunc { name, .. }
            | Self::ExportType { name, .. } | Self::ExportCoreModule { name, .. } => Some((name.as_str().into(), false)),
            Self::Import { name, ty } => Some((name.as_str().into(), matches!(ty, ComponentInterfaceRef::Instance(_)))),
            Self::ExportInstance { name, .. } => Some((name.as_str().into(), true)),
            Self::NamedImport { name, ty } => Some((name.as_ref(), matches!(ty, ComponentInterfaceRef::Instance(_)))),
            Self::NamedExport { name, kind, .. } => Some((name.as_ref(), *kind == ComponentItemKind::Instance)),
            _ => None,
        }
    }

    pub fn normalized_core_type(&self) -> Option<std::borrow::Cow<'_, ComponentCoreType>> {
        match self {
            Self::CoreType(ty) => Some(std::borrow::Cow::Borrowed(ty)),
            Self::CoreFuncType { params, results } => Some(std::borrow::Cow::Owned(
                ComponentCoreType::scalar_function(params.clone(), results.clone()))),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ComponentInterfaceRef {
    Module(u32),
    Func(u32),
    Value(ValSpec),
    ValueEq(u32),
    TypeEq(u32),
    TypeSubResource,
    Component(u32),
    Instance(u32),
}

/// Scalar-only canonical signatures and compatibility constructors.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComponentCoreValType {
    I32,
    I64,
    F32,
    F64,
    V128,
}

/// Full source grammar; indexed references resolve only in the scoped graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComponentCoreValueType {
    Scalar(ComponentCoreValType),
    Ref(ComponentCoreRefType),
}

impl ComponentCoreValueType {
    pub const I32: Self = Self::Scalar(ComponentCoreValType::I32);
    pub const I64: Self = Self::Scalar(ComponentCoreValType::I64);
    pub const F32: Self = Self::Scalar(ComponentCoreValType::F32);
    pub const F64: Self = Self::Scalar(ComponentCoreValType::F64);
    pub const V128: Self = Self::Scalar(ComponentCoreValType::V128);

    pub fn scalar(self) -> Option<ComponentCoreValType> {
        match self {
            Self::Scalar(value) => Some(value),
            Self::Ref(_) => None,
        }
    }
}

impl From<ComponentCoreValType> for ComponentCoreValueType {
    fn from(value: ComponentCoreValType) -> Self { Self::Scalar(value) }
}

/// Unresolved core source indices, never Runtime's one-based type handles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComponentCoreHeapType {
    Concrete(u32),
    Abstract(ComponentCoreAbstractHeapType),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComponentCoreAbstractHeapType {
    Func, Extern, Any, None, NoExtern, NoFunc, Eq, Struct, Array, I31,
    Exn, NoExn, Cont, NoCont,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ComponentCoreRefType {
    pub heap: ComponentCoreHeapType,
    pub nullable: bool,
    pub exact: bool,
    pub shared: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ComponentCoreStorageType {
    I8,
    I16,
    Value(ComponentCoreValueType),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComponentCoreFieldType {
    pub storage: ComponentCoreStorageType,
    pub mutable: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ComponentCoreCompositeType {
    Func { params: Vec<ComponentCoreValueType>, results: Vec<ComponentCoreValueType> },
    Struct(Vec<ComponentCoreFieldType>),
    Array(ComponentCoreFieldType),
    Cont(u32),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComponentCoreSubtype {
    pub final_: bool,
    pub supertypes: Vec<u32>,
    pub shared: bool,
    pub describes: Option<u32>,
    pub descriptor: Option<u32>,
    pub composite: ComponentCoreCompositeType,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComponentCoreRecType {
    /// Explicit singleton and empty groups retain their recursive boundary.
    pub explicit: bool,
    pub types: Vec<ComponentCoreSubtype>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ComponentCoreType {
    Rec(ComponentCoreRecType),
    Module(Vec<ComponentCoreModuleDecl>),
}

impl ComponentCoreType {
    /// Section vector rows and type-space entries are different counts.
    pub fn type_count(&self) -> Option<u32> {
        match self {
            Self::Rec(group) => u32::try_from(group.types.len()).ok(),
            Self::Module(_) => Some(1),
        }
    }

    pub fn scalar_function(params: Vec<ComponentCoreValType>, results: Vec<ComponentCoreValType>) -> Self {
        Self::function(params.into_iter().map(Into::into).collect(),
            results.into_iter().map(Into::into).collect())
    }

    pub fn function(params: Vec<ComponentCoreValueType>, results: Vec<ComponentCoreValueType>) -> Self {
        Self::Rec(ComponentCoreRecType {
            explicit: false,
            types: vec![ComponentCoreSubtype {
                final_: true, supertypes: Vec::new(), shared: false,
                describes: None, descriptor: None,
                composite: ComponentCoreCompositeType::Func { params, results },
            }],
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComponentCoreLimits {
    pub initial: u64,
    pub maximum: Option<u64>,
    pub shared: bool,
    pub index64: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ComponentCoreExternType {
    Func { type_index: u32, exact: bool },
    Table { element: ComponentCoreRefType, limits: ComponentCoreLimits },
    Memory { limits: ComponentCoreLimits, page_size_log2: Option<u32> },
    Global { value: ComponentCoreValueType, mutable: bool, shared: bool },
    Tag(u32),
}

/// Order and every binder are retained, including aliases and recursive groups.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ComponentCoreModuleDecl {
    Type(ComponentCoreType),
    AliasOuterType { count: u32, index: u32 },
    Import { module: String, name: String, ty: ComponentCoreExternType },
    Export { name: String, ty: ComponentCoreExternType },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComponentItemRef {
    pub name: String,
    pub kind: ComponentItemKind,
    pub index: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComponentItemKind {
    Module,
    Func,
    Value,
    Type,
    Component,
    Instance,
}

/// Typed exports projected into the executable component's canonical spaces.
/// These indices are resolved by normalization, not inferred from host names.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComponentInstanceExport {
    pub item: ComponentItemRef,
    pub function_type: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComponentInstanceFunctionImport {
    pub instance: u32,
    pub name: String,
    pub import: u32,
    pub expected_type: u32,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ComponentInstanceLinks {
    pub instances: Vec<Vec<ComponentInstanceExport>>,
    pub imports: Vec<ComponentInstanceFunctionImport>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComponentBinaryAsyncKind {
    Stream,
    Future,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComponentBinaryAsyncOperation {
    New,
    Read,
    Write,
    CancelRead { is_async: bool },
    CancelWrite { is_async: bool },
    DropReadable,
    DropWritable,
    Forward,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComponentBinaryCanonOption {
    Utf8,
    Utf16,
    CompactUtf16,
    Memory(u32),
    Realloc(u32),
    PostReturn(u32),
    Async,
    Callback(u32),
    CoreType(u32),
    Gc,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CoreInstanceExport {
    pub name: String,
    pub kind: CoreExportKind,
    pub index: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CoreExportKind {
    Func,
    Table,
    Memory,
    Global,
    Tag,
}

/// A component's canonical index spaces and, when fully representable, its
/// ordered binary definitions. The canonical vectors remain separate because
/// `canon lift`'s `ft` indexes the TYPE space while its row sits in the canon
/// space; declaration order does not replace either index space.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ComponentSection {
    /// Source declared a component, even if its binary form is not yet
    /// representable. Binary emission must not fall back to a core module.
    pub component_source: bool,
    /// A component-layer binary is emitted only when the source can be
    /// represented without losing any declaration or import wiring.
    pub binary: Option<Arc<ComponentBinary>>,
    /// The canon section, in canonidx order.
    pub defs: Vec<CanonDecl>,
    /// The component type index space, in typeidx order.
    pub types: Vec<TypeDecl>,
    /// The component FUNCTION index space: comp funcidx -> the canonidx of the
    /// row that DEFINES that function, or `None` when nothing here defines it.
    ///
    /// A third vector and not a field on `defs`, because it is a third index
    /// space. `canon lower`'s callee indexes THIS one; its own row sits in
    /// `defs`; its `$ft` indexes `types`. One vector serving two of those is
    /// how a row ends up pointing at itself.
    ///
    /// ⛔ `Option`, not `u32`, because of `(import "x" (func $x (type $ft)))`.
    /// An IMPORTED component function has no defining row anywhere in this
    /// component and still occupies an index in declaration order — so the
    /// slot has to exist and has to be empty. A dense vector of definitions
    /// would silently renumber every function after the import, which is the
    /// same positional-alignment trap `types` avoids the same way.
    pub funcs: Vec<Option<u32>>,
    /// Bindings installed when a child component is instantiated with function imports.
    pub func_links: Vec<ComponentFuncLink>,
    /// Instantiation bindings from typed instances. Ordinary modules carry none.
    pub instance_links: Option<Arc<ComponentInstanceLinks>>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ComponentFuncLink {
    pub import: u32,
    pub supplied: u32,
    pub expected_type: u32,
    pub supplied_type: u32,
}
