;; vybe-test: wast/wat_memory_addressing/test_load16_s_sign_extends
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 4 i32.const 0xFFFF i32.store16
          i32.const 4 i32.load16_s call $log))
