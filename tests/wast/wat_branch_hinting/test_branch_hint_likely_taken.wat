;; vybe-test: wast/wat_branch_hinting/test_branch_hint_likely_taken
;; origin: languages/wast/tests/wast/test_wat_branch_hinting.rs

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
        (@metadata.code.branch_hint "\01") i32.const 1
        if (result i32) i32.const 42 else i32.const 0 end i32.const 42 call $vybe_check_i32)
)
