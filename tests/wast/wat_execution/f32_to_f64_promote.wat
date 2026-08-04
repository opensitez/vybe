;; vybe-test: wast/wat_execution/f32_to_f64_promote
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
    f32.const 2.0
    f64.promote_f32
    i32.const 2 call $vybe_check_i32))
