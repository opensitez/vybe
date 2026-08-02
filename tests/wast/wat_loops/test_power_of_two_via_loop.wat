;; vybe-test: wast/wat_loops/test_power_of_two_via_loop
;; origin: languages/wast/tests/wast/test_wat_loops.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $i i32) (local $p i32) i32.const 1 local.set $p
        block loop local.get $i i32.const 8 i32.ge_s br_if 1
          local.get $p i32.const 2 i32.mul local.set $p
          local.get $i i32.const 1 i32.add local.set $i br 0 end end
        local.get $p call $log)
)
