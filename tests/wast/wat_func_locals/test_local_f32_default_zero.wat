;; vybe-test: wast/wat_func_locals/test_local_f32_default_zero
;; origin: languages/wast/tests/wast/test_wat_func_locals.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_f32 (param f32) (param f32)
    local.get 0
    local.get 1
    f32.ne
    if
      unreachable
    end)
  (func (export "_start") (local $x f32)
  local.get $x
  f32.const 0.0 call $vybe_check_f32
)
)
