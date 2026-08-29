//! Engine-internal slots on a JVM platform object, through the class-slot owner.
//!
//! ⛔ THE RAW `STRUCT_GET/SET, 0, key` FORM IS WHAT M3 RETIRES. Every adapter
//! here addressed its representation state by interning a string constant and
//! emitting the struct op inline, which is 57 sites the class model cannot see
//! and cannot move. `platforms/dotnet` went 418 → 0 through this same seam.
//!
//! ⚠ `Internal`, never `InstanceField`. Only the variants carrying a `class`
//! reach the canonicalising path; these keys are literal storage names
//! (`__java_io_data`, `__sb_buf`) and a canonicalised one is a DIFFERENT key,
//! which reads back `undefined` rather than failing.

use vybe_compiler::primitives::class_slots::{self, ClassSlot, PlainNames, ResolvedSlot};

/// The slot a literal storage key resolves to.
pub fn field_slot(key: &str) -> ResolvedSlot {
    class_slots::resolve(&ClassSlot::internal(key), &PlainNames)
}
