;; vybe-test: wast/wat_i64_arithmetic/test_i64_sub_pos_pos
;; origin: languages/wast/tests/wast/test_wat_i64_arithmetic.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
  (func (export "_start") i64.const 20 i64.const 10 i64.sub i64.const 10 call $vybe_check_i64)
)
