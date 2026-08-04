;; vybe-test: wast/wat_i32_bitwise/test_i32_popcnt
;; origin: languages/wast/tests/wast/test_wat_i32_bitwise.rs

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
  i32.const 0x0F0F0F0F
  i32.popcnt
  i32.const 16 call $vybe_check_i32
)
)
