;; vybe-test: wast/wat_assignment/test_swap_two_locals
;; origin: languages/wast/tests/wast/test_wat_assignment.rs

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
        (local $a i32) (local $b i32) (local $t i32)
        i32.const 3 local.set $a i32.const 9 local.set $b
        local.get $a local.set $t local.get $b local.set $a local.get $t local.set $b
        local.get $a i32.const 9 call $vybe_check_i32)
)
