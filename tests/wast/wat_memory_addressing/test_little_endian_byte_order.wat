;; vybe-test: wast/wat_memory_addressing/test_little_endian_byte_order
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
          i32.const 0 i32.const 0x04030201 i32.store
          i32.const 0 i32.load8_u i32.const 1 call $vybe_check_i32))
