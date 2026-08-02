;; vybe-test: wast/wat_control_nesting/test_br_table_selects_middle
;; origin: languages/wast/tests/wast/test_wat_control_nesting.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        block block block block
          i32.const 1 br_table 0 1 2 3
        end i32.const 100 call $log br 2
        end i32.const 200 call $log br 1
        end i32.const 300 call $log br 0
        end)
)
