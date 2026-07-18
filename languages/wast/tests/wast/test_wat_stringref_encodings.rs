//! Stringref encoding ops — WTF-16 / WTF-8 / lossy-UTF-8 new+encode, plus
//! `string.is_usv_sequence` and the GC-array new/encode variants.
//!
//! Each op's RESULT is verified: strings created from memory/arrays are checked
//! by `string.eq` against a known string or by measuring/reading back, and
//! encode ops are checked by loading the written bytes. Values follow the
//! stringref proposal (proposals/stringref/Overview.md).
use crate::wat_exec;

wat_exec! {
    // ── string.new_wtf16: decode little-endian UTF-16 code units from memory ──
    // "Hi" as WTF-16 LE = 48 00 69 00; compare against UTF-8 "Hi" (48 69).
    test_string_new_wtf16_roundtrips => { r#"
(memory 1)
(data (i32.const 0) "\48\00\69\00")
(data (i32.const 10) "\48\69")
(func (export "_start")
  i32.const 0 i32.const 2 string.new_wtf16
  i32.const 10 i32.const 2 string.new_utf8
  string.eq
  call $log)
"#, "1" },

    // ── string.encode_wtf8: write UTF-8 bytes to memory, read one back ─────────
    // "AB" (41 42); encode to offset 10; load byte 0 → 0x41 = 65.
    test_string_encode_wtf8_writes_bytes => { r#"
(memory 1)
(data (i32.const 0) "\41\42")
(func (export "_start")
  i32.const 0 i32.const 2 string.new_utf8
  i32.const 10 string.encode_wtf8   ;; returns byte count (2), dropped below
  drop
  i32.const 10 i32.load8_u
  call $log)
"#, "65" },

    // encode_wtf8 return value is the byte count.
    test_string_encode_wtf8_returns_len => { r#"
(memory 1)
(data (i32.const 0) "\41\42\43")
(func (export "_start")
  i32.const 0 i32.const 3 string.new_utf8
  i32.const 10 string.encode_wtf8
  call $log)
"#, "3" },

    // encode_lossy_utf8 behaves identically for valid strings.
    test_string_encode_lossy_utf8_writes => { r#"
(memory 1)
(data (i32.const 0) "\5A")
(func (export "_start")
  i32.const 0 i32.const 1 string.new_utf8
  i32.const 10 string.encode_lossy_utf8
  drop
  i32.const 10 i32.load8_u
  call $log)
"#, "90" },

    // ── string.is_usv_sequence: native strings are always valid USV → 1 ───────
    test_string_is_usv_sequence_true => { r#"
(memory 1)
(data (i32.const 0) "\68\69")
(func (export "_start")
  i32.const 0 i32.const 2 string.new_utf8
  string.is_usv_sequence
  call $log)
"#, "1" },

    // ── string.new_wtf8_array: build a byte array, decode → measure = 2 ───────
    test_string_new_wtf8_array => { r#"
(type $A (array (mut i8)))
(func (export "_start") (local $a (ref null $A))
  i32.const 104 i32.const 105 array.new_fixed $A 2
  local.set $a
  local.get $a i32.const 0 i32.const 2 string.new_wtf8_array
  string.measure_utf8
  call $log)
"#, "2" },

    // ── string.new_lossy_utf8_array ──────────────────────────────────────────
    test_string_new_lossy_utf8_array => { r#"
(type $A (array (mut i8)))
(func (export "_start") (local $a (ref null $A))
  i32.const 65 i32.const 66 i32.const 67 array.new_fixed $A 3
  local.set $a
  local.get $a i32.const 0 i32.const 3 string.new_lossy_utf8_array
  string.measure_utf8
  call $log)
"#, "3" },

    // ── encode_wtf8_array round-trip: string → array → string → eq ────────────
    test_string_encode_wtf8_array_roundtrips => { r#"
(type $A (array (mut i8)))
(memory 1)
(data (i32.const 0) "\48\69")
(func (export "_start") (local $a (ref null $A))
  ;; make a 2-element array, encode "Hi" into it, decode back, compare.
  i32.const 0 i32.const 0 array.new_fixed $A 2
  local.set $a
  i32.const 0 i32.const 2 string.new_utf8   ;; "Hi"
  local.get $a i32.const 0 string.encode_wtf8_array   ;; returns count
  drop
  local.get $a i32.const 0 i32.const 2 string.new_wtf8_array  ;; "Hi" again
  i32.const 0 i32.const 2 string.new_utf8
  string.eq
  call $log)
"#, "1" },

    // ── encode_lossy_utf8_array returns the byte count ───────────────────────
    test_string_encode_lossy_utf8_array_count => { r#"
(type $A (array (mut i8)))
(memory 1)
(data (i32.const 0) "\41\42\43")
(func (export "_start") (local $a (ref null $A))
  i32.const 0 i32.const 0 i32.const 0 array.new_fixed $A 3
  local.set $a
  i32.const 0 i32.const 3 string.new_utf8
  local.get $a i32.const 0 string.encode_lossy_utf8_array
  call $log)
"#, "3" },
}
