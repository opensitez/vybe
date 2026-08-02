;; vybe-test: wast/wat_local_tee/test_local_reused_across_iterations
;; origin: languages/wast/tests/wast/test_wat_local_tee.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $i i32) (local $p i32)
        i32.const 1 local.set $p i32.const 5 local.set $i
        block loop
          local.get $i i32.eqz br_if 1
          local.get $p i32.const 2 i32.mul local.set $p
          local.get $i i32.const 1 i32.sub local.set $i br 0
        end end
        local.get $p call $log)
)
