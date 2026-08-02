;; vybe-test: wast/wat_control_nesting/test_br_to_outer_block_by_depth
;; origin: languages/wast/tests/wast/test_wat_control_nesting.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        block block block i32.const 9 call $log br 2 i32.const 99 call $log end
        i32.const 88 call $log end end)
)
