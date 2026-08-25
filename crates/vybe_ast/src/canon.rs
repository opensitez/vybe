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
    pub realloc: Option<u32>,
    /// `(post-return f)` — core funcidx, `canon lift` only.
    pub post_return: Option<u32>,
    /// `(callback f)` — core funcidx, 🔀 async lift only.
    pub callback: Option<u32>,
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
#[derive(Debug, Clone, PartialEq)]
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
    },
    /// `(type $t u32)`, `(type $t (list string))` — a VALUE type.
    Value(ValSpec),
    /// A `resourcetype`, `componenttype` or `instancetype`. It occupies an
    /// index; nothing yet reads its shape beyond a resource's NAME. The string
    /// is the form's keyword so a refusal can say which one it met.
    Opaque(String),
    /// `(type $t (resource (rep i32)))` — a RESOURCE type, carrying the name it
    /// was bound under.
    ///
    /// ⛔ Separate from [`Self::Opaque`] because `own`/`borrow` need the NAME,
    /// not the keyword. `component::ValType::Own` is `Own(String)` — a handle
    /// binds to whatever is registered under that name — while the source
    /// names the resource by INDEX. This is the only record of which name that
    /// index belongs to, so folding it back into `Opaque("resource")` (which
    /// is what it used to be) makes every `(own $r)` unresolvable.
    ///
    /// An UNNAMED resource type keeps `None`: it occupies its index, and a
    /// handle to it refuses rather than being given an invented name.
    Resource(Option<String>),
}

/// A component's declaration sections: the canon section and the type space.
///
/// Two `Vec`s and not one, because they are two index spaces. `canon lift`'s
/// `ft` indexes the TYPE space while its own row sits in the canon space, and
/// one vector serving both is how a row ends up pointing at itself.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ComponentSection {
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
}
