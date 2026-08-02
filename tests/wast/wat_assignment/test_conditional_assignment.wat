;; vybe-test: wast/wat_assignment/test_conditional_assignment
;; origin: languages/wast/tests/wast/test_wat_assignment.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $x i32) i32.const 1
        if i32.const 100 local.set $x else i32.const 200 local.set $x end
        local.get $x call $log)
)
