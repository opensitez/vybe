;; vybe-test: wast/wat_local_tee/test_local_f64_default_zero
;; origin: languages/wast/tests/wast/test_wat_local_tee.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $x f64) local.get $x call $log_f64)
)
