;; vybe-test: wast/wat_control_nesting/test_br_table_default_target
;; origin: languages/wast/tests/wast/test_wat_control_nesting.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        block block
          i32.const 9 br_table 0 1
        end i32.const 111 call $log br 1
        end i32.const 222 call $log)
)
