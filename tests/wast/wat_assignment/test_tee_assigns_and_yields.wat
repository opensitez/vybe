;; vybe-test: wast/wat_assignment/test_tee_assigns_and_yields
;; origin: languages/wast/tests/wast/test_wat_assignment.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $x i32) i32.const 7 local.tee $x local.get $x i32.add call $log)
)
