//! # gc proposal
//!
//! Spec: `proposals/gc/proposals/gc/MVP.md`. Adds garbage-collected reference
//! types + casts + struct/array ops. Prefix `0xFB`, plus four
//! core-prefix extensions (`ref.eq`, `ref.as_non_null`, `br_on_null`,
//! `br_on_non_null`).
//!
//! ## Byte-value compliance
//!
//! Every opcode this module declares matches the MVP binary-format
//! table exactly. `tests/wasm_compliance_test.rs::gc_opcodes_use_spec_byte_values`
//! pins the full table as a regression guard.
//!
//! Historical bugs fixed:
//!   * `REF_EQ` used to live at `0xFB 0x13` — which is `array.init_elem`
//!     in the spec. Moved to core `0x00 0xD3`.
//!   * `ARRAY_NEW` used to encode at `0xFB 0x06` but implement the
//!     N-value `array.new_fixed` semantics. Split into a proper
//!     `ARRAY_NEW` (`0xFB 0x06`, single-value-and-length) and
//!     `ARRAY_NEW_FIXED` (`0xFB 0x08`, N stacked values).
//!
//! ## Status
//!
//! | Feature                            | Status | Byte       | Notes                                           |
//! |------------------------------------|--------|------------|-------------------------------------------------|
//! | `struct.new`                       | ✅     | `fb 00`    |                                                 |
//! | `struct.new_default`               | ✅     | `fb 01`    | zero-inits fields to null                       |
//! | `struct.get`                       | ✅     | `fb 02`    | field-indexed                                   |
//! | `struct.get_s` / `struct.get_u`    | ✅     | `fb 03/04` | identity in our externref-only struct model     |
//! | `struct.set`                       | ✅     | `fb 05`    |                                                 |
//! | `array.new`                        | ✅     | `fb 06`    | `[value, length] → array`                       |
//! | `array.new_default`                | ✅     | `fb 07`    | `[length] → null-filled array`                  |
//! | `array.new_fixed`                  | ✅     | `fb 08`    | `[v1..vN] → array`                              |
//! | `array.new_data` / `new_elem`      | ⚠     | `fb 09/0a` | emitter spec-correct; VM stubs (no data segs)   |
//! | `array.get` / `set` / `len` / `fill` / `copy` | ✅ | `fb 0b–11` |                                              |
//! | `array.get_s` / `array.get_u`      | ✅     | `fb 0c/0d` | identity in our externref-only array model      |
//! | `array.init_data` / `init_elem`    | ⚠     | `fb 12/13` | emitter spec-correct; VM stubs                  |
//! | `ref.test` / `ref.test_null`       | ✅     | `fb 14/15` |                                                 |
//! | `ref.cast` / `ref.cast_null`       | ✅     | `fb 16/17` |                                                 |
//! | `br_on_cast` / `br_on_cast_fail`   | ✅     | `fb 18/19` |                                                 |
//! | `any.convert_extern` / `extern.convert_any` | ✅ | `fb 1a/1b` | identity over our universal externref    |
//! | `ref.i31` / `i31.get_s` / `i31.get_u`       | ✅ | `fb 1c–1e` |                                          |
//! | `ref.eq`                           | ✅     | `00 d3`    | core prefix                                     |
//! | `ref.as_non_null`                  | ✅     | `00 d4`    | traps on null per spec                          |
//! | `br_on_null` / `br_on_non_null`    | ✅     | `00 d5/d6` |                                                 |
//! | Shared GC objects                  | ❌     |            | needs atomic struct access; deferred            |
//!
//! ## Emitter
//!
//! The 0xFB-prefix emitter lives in `code.rs::emit_gc_op` (access to
//! type context needed for struct/array type indices). The four core
//! extensions (`ref.eq`, `ref.as_non_null`, `br_on_null`,
//! `br_on_non_null`) emit from `emit_core_op`.

use crate::Chunk;

/// GC proposal declares no imports — types live in the type section.
pub const IMPORTS: &[(&str, &str)] = &[];
/// No globals either.
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];

pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}
