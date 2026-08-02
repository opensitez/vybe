;; vybe-test: wast/wat_i64_relational/test_i64_eqz
;; origin: languages/wast/tests/wast/test_wat_i64_relational.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i64.const 0
  i64.eqz
  call $log
)
)
