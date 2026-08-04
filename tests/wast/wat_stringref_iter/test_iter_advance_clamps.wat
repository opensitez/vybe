;; vybe-test: wast/wat_stringref_iter/test_iter_advance_clamps
;; origin: languages/wast/tests/wast/test_wat_stringref_iter.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start")
  i32.const 0 i32.const 5 string.new_utf8
  string.as_iter
  i32.const 99
  stringview_iter.advance
  i32.const 5 call $vybe_check_i32)
)
