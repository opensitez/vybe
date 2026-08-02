;; vybe-test: wast/wat_assignment/test_reassignment_overwrites
;; origin: languages/wast/tests/wast/test_wat_assignment.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $x i32) i32.const 1 local.set $x i32.const 2 local.set $x
        local.get $x call $log)
)
