;; vybe-test: wast/wat_memory_addressing/test_load8_s_sign_extends
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 200 i32.store8
          i32.const 0 i32.load8_s i32.const -56 call $vybe_check_i32))
