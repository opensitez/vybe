;; vybe-test: wast/wat_control_nesting/test_nested_loop_accumulate
;; origin: languages/wast/tests/wast/test_wat_control_nesting.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $sum i32) (local $i i32)
        i32.const 5 local.set $i
        block loop
          local.get $i i32.eqz br_if 1
          local.get $sum local.get $i i32.add local.set $sum
          local.get $i i32.const 1 i32.sub local.set $i
          br 0
        end end
        local.get $sum call $log)
)
