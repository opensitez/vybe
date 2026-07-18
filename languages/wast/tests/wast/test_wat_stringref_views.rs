//! Stringview cursor ops (WTF-8 byte offsets, WTF-16 code-unit offsets).
//! `string.as_wtf8`/`as_wtf16` yield the view; the ops below index/slice/encode
//! it with explicit position operands. Results verified by measuring/comparing
//! the produced string or reading back the encoded bytes. Semantics per
//! proposals/stringref/Overview.md (the "position treatment" clamps/snaps).
use crate::wat_exec;

wat_exec! {
    // ── stringview_wtf16.length: number of UTF-16 code units ("Hi" → 2) ───────
    test_wtf16_length => { r#"
(memory 1)
(data (i32.const 0) "\48\69")
(func (export "_start")
  i32.const 0 i32.const 2 string.new_utf8
  string.as_wtf16
  stringview_wtf16.length
  call $log)
"#, "2" },

    // ── stringview_wtf16.get_codeunit: unit at pos 1 of "Hi" is 'i' = 105 ─────
    test_wtf16_get_codeunit => { r#"
(memory 1)
(data (i32.const 0) "\48\69")
(func (export "_start")
  i32.const 0 i32.const 2 string.new_utf8
  string.as_wtf16
  i32.const 1
  stringview_wtf16.get_codeunit
  call $log)
"#, "105" },

    // ── stringview_wtf16.slice: "Hello"[1..3] = "el", measure = 2 ─────────────
    test_wtf16_slice => { r#"
(memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start")
  i32.const 0 i32.const 5 string.new_utf8
  string.as_wtf16
  i32.const 1 i32.const 3
  stringview_wtf16.slice
  string.measure_utf8
  call $log)
"#, "2" },

    // slice content check: "Hello"[1..3] equals "el".
    test_wtf16_slice_content => { r#"
(memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(data (i32.const 20) "\65\6C")
(func (export "_start")
  i32.const 0 i32.const 5 string.new_utf8
  string.as_wtf16
  i32.const 1 i32.const 3 stringview_wtf16.slice
  i32.const 20 i32.const 2 string.new_utf8
  string.eq
  call $log)
"#, "1" },

    // ── stringview_wtf16.encode: write 2 code units, return count 2 ───────────
    test_wtf16_encode_count => { r#"
(memory 1)
(data (i32.const 0) "\48\69")
(func (export "_start")
  i32.const 0 i32.const 2 string.new_utf8
  string.as_wtf16
  i32.const 10 i32.const 0 i32.const 2
  stringview_wtf16.encode
  call $log)
"#, "2" },

    // encode writes little-endian u16: first byte of 'H'(0x48) at ptr 10 = 72.
    test_wtf16_encode_bytes => { r#"
(memory 1)
(data (i32.const 0) "\48\69")
(func (export "_start")
  i32.const 0 i32.const 2 string.new_utf8
  string.as_wtf16
  i32.const 10 i32.const 0 i32.const 2 stringview_wtf16.encode
  drop
  i32.const 10 i32.load16_u
  call $log)
"#, "72" },

    // ── stringview_wtf8.advance: from 0 by 2 bytes over "Hello" → 2 ───────────
    test_wtf8_advance => { r#"
(memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start")
  i32.const 0 i32.const 5 string.new_utf8
  string.as_wtf8
  i32.const 0 i32.const 2
  stringview_wtf8.advance
  call $log)
"#, "2" },

    // advance past the end clamps to the byte length (5).
    test_wtf8_advance_clamps => { r#"
(memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start")
  i32.const 0 i32.const 5 string.new_utf8
  string.as_wtf8
  i32.const 0 i32.const 99
  stringview_wtf8.advance
  call $log)
"#, "5" },

    // ── stringview_wtf8.slice: "Hello"[1..3] = "el", measure = 2 ──────────────
    test_wtf8_slice => { r#"
(memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start")
  i32.const 0 i32.const 5 string.new_utf8
  string.as_wtf8
  i32.const 1 i32.const 3
  stringview_wtf8.slice
  string.measure_utf8
  call $log)
"#, "2" },
}
