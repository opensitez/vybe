;; vybe-test: wast/wat_string_new/test_string_new_lossy_utf8
;; origin: languages/wast/tests/wast/test_wat_string_new.rs

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
(data (i32.const 0) "\ff\ff") ;; invalid utf8
(func (export "_start")
  i32.const 0
  i32.const 2
  string.new_lossy_utf8
  string.measure_utf8
  i32.const 6 call $vybe_check_i32
)
)
