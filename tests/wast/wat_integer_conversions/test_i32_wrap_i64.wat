;; vybe-test: wast/wat_integer_conversions/test_i32_wrap_i64
;; origin: languages/wast/tests/wast/test_wat_integer_conversions.rs

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
  i64.const 4294967338 ;; 0x10000002A -> wraps to 42
  i32.wrap_i64
  i32.const 42 call $vybe_check_i32
)
)
