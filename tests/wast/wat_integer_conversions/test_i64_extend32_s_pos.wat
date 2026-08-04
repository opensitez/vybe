;; vybe-test: wast/wat_integer_conversions/test_i64_extend32_s_pos
;; origin: languages/wast/tests/wast/test_wat_integer_conversions.rs

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
  (func (export "_start")
  i64.const 2147483647
  i64.extend32_s
  i64.const 2147483647 call $vybe_check_i64
)
)
