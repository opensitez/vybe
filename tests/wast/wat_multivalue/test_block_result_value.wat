;; vybe-test: wast/wat_multivalue/test_block_result_value
;; origin: languages/wast/tests/wast/test_wat_multivalue.rs

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
        block (result i32) i32.const 5 i32.const 6 i32.add end i32.const 11 call $vybe_check_i32)
)
