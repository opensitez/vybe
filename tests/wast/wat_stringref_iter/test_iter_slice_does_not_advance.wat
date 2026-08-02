;; vybe-test: wast/wat_stringref_iter/test_iter_slice_does_not_advance
;; origin: languages/wast/tests/wast/test_wat_stringref_iter.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
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
)
