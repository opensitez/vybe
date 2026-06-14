//! # stack-switching proposal
//!
//! Spec: WebAssembly/stack-switching. Adds first-class coroutines /
//! continuations via six instructions and a new composite type.
//!
//! Binary encoding (core prefix `0x00`, subs `0xE0..=0xE6`):
//!
//! | Op                | Sub    | Immediates                                     |
//! |-------------------|--------|------------------------------------------------|
//! | `cont.new`        | `0xE0` | `typeidx` (continuation type)                  |
//! | `cont.bind`       | `0xE1` | `src_typeidx`, `dst_typeidx`                   |
//! | `suspend`         | `0xE2` | `tagidx`                                       |
//! | `resume`          | `0xE3` | `cont_typeidx`, `vec(handler)`                 |
//! | `resume_throw`    | `0xE4` | `cont_typeidx`, `tagidx`, `vec(handler)`       |
//! | `resume_throw_ref`| `0xE5` | `cont_typeidx`, `vec(handler)`                 |
//! | `switch`          | `0xE6` | `cont_typeidx`, `tagidx`                        |
//!
//! A `handler` is `0x00 tagidx labelidx` (on-tag-to-label) or
//! `0x01 tagidx` (on-tag-to-switch).
//!
//! Continuation types live in the type section as `0x5D <funcidx>` —
//! a `(cont $ft)` wraps a function-type index whose signature matches
//! the fiber's entry function. We emit one continuation type per
//! distinct function arity observed in the bytecode's `CONT_NEW` ops.
//!
//! Tags for suspend/resume carry `(param T*) (result R*)` signatures —
//! the yield-type and resume-type of the coroutine. We declare a
//! single "vybe.suspend" tag with `(param externref) (result externref)`
//! as the generic shape; typed-continuation flows can introduce more.
//!
//! ## Status in Vybe
//!
//! | Feature                     | Status |
//! |-----------------------------|--------|
//! | `cont.new $ct`              | ✅ `Op::CONT_NEW` emits `0xE0 <cont_typeidx>` |
//! | `suspend $tag`              | ✅ `Op::SUSPEND` emits `0xE2 <tagidx=1>` |
//! | `resume $ct handlers`       | ✅ `Op::RESUME` emits `0xE3 <cont_typeidx> handlers` and VM routes matching suspend tags |
//! | `switch $ct $tag`           | ✅ `Op::SWITCH` emits `0xE6 <cont_typeidx> <tagidx=1>` |
//! | `cont.bind $ct1 $ct2`       | ✅ `Op::CONT_BIND` emits `0xE1 <src_ct> <dst_ct>` |
//! | `resume_throw $ct $tag h`   | ✅ `Op::RESUME_THROW` emits `0xE4 <cont_ct> <tagidx> handlers` |
//! | `resume_throw_ref $ct h`    | ✅ `Op::RESUME_THROW_REF` emits `0xE5 <cont_ct> handlers` |
//! | Continuation type `(cont $ft)` in type section | ✅ emitted as `0x5D <funcidx>` when any stack-switching op is present |
//! | Suspend tag in tag section  | ✅ declared alongside the exception tag when stack-switching is used |
//! | VM coroutine semantics      | ✅ `ObjectKind::Continuation` holds entry + `Fiber`; `SUSPEND` saves via `save_fiber`, `RESUME` restores via `resume_fiber_with`; `active_continuations` stack routes suspend to the right cont |
//! | Typed continuations         | ✅ `SUSPEND_TYPED` emits a distinct Wasm tag for each declared continuation tag; VM checks typed `cont.new` / `resume` contracts |
//!
//! Untyped stack-switch ops still reference the generic `vybe.suspend`
//! tag; typed suspend paths use the per-tag entries declared after it
//! in the tag section.

use crate::Chunk;

/// Reference-types declares no globals.
pub fn declare_imports() -> &'static [(&'static str, &'static str)] {
    &[]
}
pub fn declare_globals() -> &'static [(&'static str, &'static str)] {
    &[]
}
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}

// ── Binary encoding constants ───────────────────────────────────────

pub const CONT_TYPE_PREFIX: u8 = 0x5D; // (cont $ft) in the type section
pub const OP_CONT_NEW: u8 = 0xE0;
pub const OP_CONT_BIND: u8 = 0xE1;
pub const OP_SUSPEND: u8 = 0xE2;
pub const OP_RESUME: u8 = 0xE3;
pub const OP_RESUME_THROW: u8 = 0xE4;
pub const OP_RESUME_THROW_REF: u8 = 0xE5;
pub const OP_SWITCH: u8 = 0xE6;
