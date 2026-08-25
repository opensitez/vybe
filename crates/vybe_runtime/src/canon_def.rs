//! The canon definition section — `Binary.md` §"Canonical Definitions".
//!
//! `Binary.md:296` is not a list of opcodes; it is a section of **typed
//! definition records**, and almost every row carries immediates:
//!
//! ```text
//! 0x00 0x00 f:<core:funcidx> opts:<opts> ft:<typeidx>  (canon lift  f opts ft)
//! 0x01 0x00 f:<funcidx>      opts:<opts>               (canon lower f opts)
//! 0x0f t:<typeidx> opts:<opts>                         (canon stream.read  t opts)
//! 0x09 rs:<resultlist> opts:<opts>                     (canon task.return rs opts)
//! 0x0a v:<valtype> i:<u32>                             (canon context.get v i)
//! 0x0c cancel?:<cancel?>                               (canon thread.yield cancel?)
//! 0x27 ft:<core:typeidx> tbl:<core:tableidx>           (canon thread.new-indirect ft tbl)
//! ```
//!
//! Before this existed the VM carried **one** `Option<u32>`, smuggled through
//! the import name as `@N`. Rows that were implemented and shipped therefore
//! ran with immediates that had nowhere to live: a `stream<string>` could not
//! say utf8 vs utf16, `task.return` could not check its result type, and every
//! `cancel?` / `async?` was silently absent.
//!
//! These are **instantiation-time** immediates in the spec — `Store.lift` /
//! `Store.lower` capture them once and produce a `FuncInst`. A canon import
//! resolving to a row of this table at link time is that same capture: the
//! index is what makes two instantiations of one built-in distinct core funcs.

use crate::component::ValType;

/// `string-encoding` — `canonopt` 0x00/0x01/0x02.
///
/// "When there is no `string-encoding` present, the default value is `utf8`"
/// (`Binary.md`), which is why this has a `Default` and the others do not.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum StringEncoding {
    #[default]
    Utf8,
    Utf16,
    Latin1Utf16,
}

/// `opts ::= opt*:vec(<canonopt>)`.
///
/// `memory` and `realloc` are kept even though this VM has exactly one of each
/// (`SharedMemory`, `bump_realloc`). The resolver ASSERTS the immediate names
/// that one rather than dropping the field — same reasoning as
/// `canon_element_type` refusing instead of defaulting: on the day a second
/// memory exists, a stale index must fail loudly instead of aliasing.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CanonOpts {
    pub string_encoding: StringEncoding,
    /// `(memory m)` — core memidx.
    pub memory: Option<u32>,
    /// `(realloc f)` — core funcidx.
    pub realloc: Option<u32>,
    /// `(post-return f)` — core funcidx, `canon lift` only.
    pub post_return: Option<u32>,
    /// `(callback f)` — core funcidx, 🔀 async lift only.
    pub callback: Option<u32>,
    /// `async` 🔀 — selects `MAX_FLAT_ASYNC_PARAMS` over `MAX_FLAT_PARAMS`.
    pub is_async: bool,
}

/// A Component Model function type: N params, at most ONE result.
///
/// Not `component::FuncSig`, which carries `results: Vec<ValType>` for the host
/// interface surface. The canonical ABI is single-result — `MAX_FLAT_RESULTS`
/// is 1 and `flatten_functype` takes `Option<&ValType>` — so a `Vec` here would
/// admit a shape the ABI cannot express and push the error to the first lower.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CanonFuncType {
    pub params: Vec<ValType>,
    pub result: Option<ValType>,
}

/// What a `canon lift` / `canon lower` wraps.
///
/// The sort is carried, not inferred. `canon lift $callee` takes a **core**
/// funcidx; `canon lower $callee` takes a **component** funcidx. A bare `u32`
/// makes those indistinguishable, and a mis-wired definition would then lift
/// with core types and look entirely correct. The discriminator that proves the
/// wiring: a `lower` callee must resolve to something with a `ValType`
/// signature (it has to know what to lift flat args INTO), a `lift` callee to a
/// flat core signature.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CalleeRef {
    /// `core:funcidx` — the callee of `canon lift`.
    Core(u32),
    /// `funcidx` in the component's function index space — `canon lower`.
    Component(u32),
}

/// One row of the canon section.
///
/// Every immediate any in-scope row can carry, in one record. Rows that take
/// none leave the fields at their defaults; a row that needs one and finds
/// `None` refuses rather than guessing, because a canonical ABI that is quietly
/// wrong about layout corrupts a peer's memory with nothing to detect it.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CanonDef {
    /// `$callee` — `lift` / `lower` only.
    pub callee: Option<CalleeRef>,
    /// `t:<typeidx>` (stream/future element), `rt:<typeidx>` (resource), or
    /// `ft:<typeidx>` for `lift` — an index into the component type space.
    pub ty: Option<u32>,
    /// `ft` as a FUNCTION type, for `canon lift`. Separate from `ty` because
    /// they index different spaces: a value type and a function type.
    pub functype: Option<u32>,
    /// `rs:<resultlist>` — `canon task.return`.
    pub results: Vec<u32>,
    /// `(v, i)` — `canon context.get` / `context.set`. Two immediates, which is
    /// precisely what a single `Option<u32>` channel could never carry.
    pub context: Option<(ValType, u32)>,
    /// `tbl:<core:tableidx>` — `thread.new-indirect`, `thread.spawn-indirect`.
    pub table: Option<u32>,
    pub opts: CanonOpts,
    /// `cancel?` — `thread.{suspend,yield,...}`, `{stream,future}.cancel-*`.
    pub cancellable: bool,
    /// `async?` — `subtask.cancel`, `{stream,future}.cancel-*`.
    pub is_async: bool,
    /// `shared?` 🧵② — the spawn rows.
    pub shared: bool,
}

impl CanonDef {
    /// A row carrying only an element/resource type — `stream.new`,
    /// `future.new`, `resource.{new,rep,drop}`, `stream.drop-*`.
    pub fn with_type(ty: u32) -> Self {
        CanonDef {
            ty: Some(ty),
            ..Default::default()
        }
    }

    /// A row carrying an element type plus options — `stream.read`,
    /// `stream.write`, `future.read`, `future.write`.
    pub fn with_type_and_opts(ty: u32, opts: CanonOpts) -> Self {
        CanonDef {
            ty: Some(ty),
            opts,
            ..Default::default()
        }
    }

    /// `cancel?`-only rows — `thread.yield`, `waitable-set.{wait,poll}`.
    pub fn cancellable(cancellable: bool) -> Self {
        CanonDef {
            cancellable,
            ..Default::default()
        }
    }

    /// The element type this row declares, or an error naming the row.
    /// Refuses rather than defaulting — substituting a plausible width moves
    /// the wrong number of bytes and nothing downstream can detect it.
    pub fn require_type(&self, builtin: &str) -> Result<u32, String> {
        self.ty.ok_or_else(|| {
            format!("canon {builtin}: definition carries no $t immediate")
        })
    }

    /// The callee this row wraps, or an error. `lift`/`lower` only.
    pub fn require_callee(&self, builtin: &str) -> Result<CalleeRef, String> {
        self.callee.ok_or_else(|| {
            format!("canon {builtin}: definition carries no $callee immediate")
        })
    }
}

/// The module's canon section: `canonidx -> CanonDef`.
///
/// Shared by every chunk exactly as `Chunk::globals` is, because the section is
/// module-level in the binary format — a chunk that CARRIES a canonidx must be
/// able to say what that index means, and threading the table through every
/// caller is what dragged `cli.rs` in last time.
pub type CanonSection = std::sync::Arc<Vec<CanonDef>>;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn string_encoding_defaults_to_utf8() {
        // Binary.md: "When there is no `string-encoding` present, the default
        // value is `utf8`. For all other values, the default value is 'none'."
        let o = CanonOpts::default();
        assert_eq!(o.string_encoding, StringEncoding::Utf8);
        assert_eq!(o.memory, None, "no memory is 'none', not memory 0");
        assert_eq!(o.realloc, None);
        assert_eq!(o.post_return, None);
        assert_eq!(o.callback, None);
        assert!(!o.is_async);
    }

    #[test]
    fn a_row_refuses_a_type_it_was_never_given() {
        // The whole reason the table exists: an absent immediate must be an
        // error naming the missing declaration, never a guessed width.
        let d = CanonDef::default();
        let err = d.require_type("stream.read").unwrap_err();
        assert!(err.contains("stream.read"), "the error names the row: {err}");
        assert!(err.contains("$t"));
    }

    #[test]
    fn a_callee_carries_its_sort() {
        // `lift` takes a CORE funcidx, `lower` a COMPONENT funcidx. Same
        // integer, different index spaces — a bare u32 would make a mis-wired
        // definition look correct.
        let lift = CanonDef {
            callee: Some(CalleeRef::Core(3)),
            ..Default::default()
        };
        let lower = CanonDef {
            callee: Some(CalleeRef::Component(3)),
            ..Default::default()
        };
        assert_ne!(lift.callee, lower.callee, "3 core != 3 component");
        assert_eq!(lift.require_callee("lift").unwrap(), CalleeRef::Core(3));
    }

    #[test]
    fn context_carries_two_immediates() {
        // `canon context.get v i` — the case that a single Option<u32> could
        // not express at all.
        let d = CanonDef {
            context: Some((ValType::I32, 1)),
            ..Default::default()
        };
        let (t, i) = d.context.clone().unwrap();
        assert_eq!(t, ValType::I32);
        assert_eq!(i, 1, "Explainer.md restricts i < 2");
    }

    #[test]
    fn ty_and_functype_are_different_index_spaces() {
        // `canon lift`'s `ft` is a FUNCTION type; `stream.read`'s `t` is a
        // value type. Conflating them is the `GLOBAL_GET` defect — one integer
        // serving two spaces.
        let d = CanonDef {
            ty: Some(0),
            functype: Some(0),
            ..Default::default()
        };
        assert_eq!(d.ty, d.functype, "same integer");
        // ...and yet they mean different things, which is why there are two
        // fields rather than one.
        assert!(d.ty.is_some() && d.functype.is_some());
    }

    #[test]
    fn a_canon_functype_has_at_most_one_result() {
        // MAX_FLAT_RESULTS is 1 and `flatten_functype` takes Option<&ValType>.
        // A Vec here would admit a shape the ABI cannot express.
        let ft = CanonFuncType {
            params: vec![ValType::I32, ValType::String],
            result: Some(ValType::Bool),
        };
        assert_eq!(ft.params.len(), 2);
        assert!(ft.result.is_some());
        assert_eq!(CanonFuncType::default().result, None, "a void function");
    }
}
