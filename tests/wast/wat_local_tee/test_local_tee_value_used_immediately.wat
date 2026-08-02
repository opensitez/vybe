;; vybe-test: wast/wat_local_tee/test_local_tee_value_used_immediately
;; origin: languages/wast/tests/wast/test_wat_local_tee.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $x i32) i32.const 5 local.tee $x i32.const 3 i32.add call $log)
)
