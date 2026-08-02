;; vybe-test: wast/wat_branch_hinting/test_branch_hint_on_br_if
;; origin: languages/wast/tests/wast/test_wat_branch_hinting.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        block i32.const 7 call $log
          (@metadata.code.branch_hint "\01") i32.const 1 br_if 0
          i32.const 8 call $log
        end)
)
