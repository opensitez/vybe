;; vybe-test: wast/wat_conditionals/test_boolean_and
;; origin: languages/wast/tests/wast/test_wat_conditionals.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i32.const 1 i32.const 1 i32.and call $log)
)
