;; vybe-test: wast/wat_stringref_iter/test_iter_rewind_clamps
;; origin: languages/wast/tests/wast/test_wat_stringref_iter.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
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
)
