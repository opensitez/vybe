;; vybe-test: wast/wat_string_new/test_string_new_utf8_oob
;; origin: languages/wast/tests/wast/test_wat_string_new.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(func (export "_start")
  i32.const 65530
  i32.const 10
  string.new_utf8
  drop
  i32.const 42
  call $log
)
)
