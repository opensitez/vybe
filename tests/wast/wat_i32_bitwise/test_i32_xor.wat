;; vybe-test: wast/wat_i32_bitwise/test_i32_xor
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
  i32.const 0xFF
  i32.const 0xAA
  i32.xor
  i32.const 85 call $vybe_check_i32
)
)
