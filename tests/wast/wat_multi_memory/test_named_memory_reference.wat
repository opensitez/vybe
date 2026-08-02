;; vybe-test: wast/wat_multi_memory/test_named_memory_reference
;; origin: languages/wast/tests/wast/test_wat_multi_memory.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory $data 1) (memory $scratch 1)
        (func (export "_start")
          i32.const 4 i32.const 999 i32.store $scratch
          i32.const 4 i32.load $scratch call $log))
