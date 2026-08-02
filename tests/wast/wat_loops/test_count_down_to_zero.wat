;; vybe-test: wast/wat_loops/test_count_down_to_zero
;; origin: languages/wast/tests/wast/test_wat_loops.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $i i32) i32.const 7 local.set $i
        block loop local.get $i i32.eqz br_if 1
          local.get $i i32.const 1 i32.sub local.set $i br 0 end end
        local.get $i call $log)
)
