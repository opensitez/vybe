;; vybe-test: wast/wat_if_else/test_if_break
;; origin: languages/wast/tests/wast/test_wat_if_else.rs

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
  block (result i32)
    i32.const 1
    if (result i32)
      i32.const 42
      br 1
    else
      i32.const 99
    end
  end
  i32.const 42 call $vybe_check_i32
)
)
