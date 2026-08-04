;; vybe-test: wast/wat_control_nesting/test_nested_loop_accumulate
;; origin: languages/wast/tests/wast/test_wat_control_nesting.rs

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
  (func (export "_start")
        (local $sum i32) (local $i i32)
        i32.const 5 local.set $i
        block loop
          local.get $i i32.eqz br_if 1
          local.get $sum local.get $i i32.add local.set $sum
          local.get $i i32.const 1 i32.sub local.set $i
          br 0
        end end
        local.get $sum i32.const 15 call $vybe_check_i32)
)
