//! Codepoint-iterator stringview ops: `string.as_iter` yields a cursor;
//! `stringview_iter.next` returns the current codepoint (or -1) and advances,
//! `advance`/`rewind` move the cursor and return the count actually consumed,
//! `slice` returns up to N codepoints from the cursor without advancing.
//! Semantics per proposals/stringref/Overview.md.
use crate::wat_exec;

wat_exec! {
    // ── next: first codepoint of "Hi" is 'H' = 72 ────────────────────────────
    test_iter_next_first => { r#"
(memory 1)
(data (i32.const 0) "\48\69")
(func (export "_start")
  i32.const 0 i32.const 2 string.new_utf8
  string.as_iter
  stringview_iter.next
  call $log)
"#, "72" },

    // next advances: two calls give 'H'(72) then 'i'(105); second value logged.
    test_iter_next_advances => { r#"
(memory 1)
(data (i32.const 0) "\48\69")
(func (export "_start") (local $it (ref null $dummy))
  i32.const 0 i32.const 2 string.new_utf8
  string.as_iter
  local.set $it
  local.get $it stringview_iter.next drop   ;; consume 'H'
  local.get $it stringview_iter.next        ;; 'i' = 105
  call $log)
(type $dummy (struct))
"#, "105" },

    // next at end returns -1.
    test_iter_next_at_end => { r#"
(memory 1)
(data (i32.const 0) "\41")
(func (export "_start") (local $it (ref null $dummy))
  i32.const 0 i32.const 1 string.new_utf8
  string.as_iter
  local.set $it
  local.get $it stringview_iter.next drop   ;; consume 'A'
  local.get $it stringview_iter.next        ;; end → -1
  call $log)
(type $dummy (struct))
"#, "-1" },

    // ── advance returns the number of codepoints actually consumed ────────────
    test_iter_advance_count => { r#"
(memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start")
  i32.const 0 i32.const 5 string.new_utf8
  string.as_iter
  i32.const 3
  stringview_iter.advance
  call $log)
"#, "3" },

    // advance past the end consumes only what's left (5 for a 5-char string).
    test_iter_advance_clamps => { r#"
(memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start")
  i32.const 0 i32.const 5 string.new_utf8
  string.as_iter
  i32.const 99
  stringview_iter.advance
  call $log)
"#, "5" },

    // ── rewind after advancing returns the consumed count ─────────────────────
    test_iter_rewind_count => { r#"
(memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start") (local $it (ref null $dummy))
  i32.const 0 i32.const 5 string.new_utf8
  string.as_iter
  local.set $it
  local.get $it i32.const 4 stringview_iter.advance drop  ;; at pos 4
  local.get $it i32.const 2 stringview_iter.rewind        ;; back 2
  call $log)
(type $dummy (struct))
"#, "2" },

    // rewind past the start only rewinds to 0.
    test_iter_rewind_clamps => { r#"
(memory 1)
(data (i32.const 0) "\48\65")
(func (export "_start") (local $it (ref null $dummy))
  i32.const 0 i32.const 2 string.new_utf8
  string.as_iter
  local.set $it
  local.get $it i32.const 1 stringview_iter.advance drop  ;; pos 1
  local.get $it i32.const 9 stringview_iter.rewind        ;; only 1 back
  call $log)
(type $dummy (struct))
"#, "1" },

    // ── slice returns up to N codepoints from the cursor, without advancing ───
    // "Hello", advance 1 (cursor at 'e'), slice 3 → "ell", measure = 3.
    test_iter_slice => { r#"
(memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start") (local $it (ref null $dummy))
  i32.const 0 i32.const 5 string.new_utf8
  string.as_iter
  local.set $it
  local.get $it i32.const 1 stringview_iter.advance drop
  local.get $it i32.const 3 stringview_iter.slice
  string.measure_utf8
  call $log)
(type $dummy (struct))
"#, "3" },

    // slice does NOT advance: next after slice still returns the cursor char.
    test_iter_slice_does_not_advance => { r#"
(memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start") (local $it (ref null $dummy))
  i32.const 0 i32.const 5 string.new_utf8
  string.as_iter
  local.set $it
  local.get $it i32.const 2 stringview_iter.slice drop  ;; slice, no advance
  local.get $it stringview_iter.next                    ;; still 'H' = 72
  call $log)
(type $dummy (struct))
"#, "72" },
}
