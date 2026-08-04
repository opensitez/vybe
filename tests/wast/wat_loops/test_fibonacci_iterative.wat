;; vybe-test: wast/wat_loops/test_fibonacci_iterative
;; origin: languages/wast/tests/wast/test_wat_loops.rs

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
        (local $a i32) (local $b i32) (local $i i32) (local $t i32)
        i32.const 0 local.set $a i32.const 1 local.set $b
        block loop
          local.get $i i32.const 10 i32.ge_s br_if 1
          local.get $a local.get $b i32.add local.set $t
          local.get $b local.set $a local.get $t local.set $b
          local.get $i i32.const 1 i32.add local.set $i br 0
        end end local.get $a i32.const 55 call $vybe_check_i32)
)
