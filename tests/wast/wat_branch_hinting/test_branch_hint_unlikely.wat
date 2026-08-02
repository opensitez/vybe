;; vybe-test: wast/wat_branch_hinting/test_branch_hint_unlikely
;; origin: languages/wast/tests/wast/test_wat_branch_hinting.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (@metadata.code.branch_hint "\00") i32.const 0
        if (result i32) i32.const 1 else i32.const 99 end call $log)
)
