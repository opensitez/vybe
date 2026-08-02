;; vybe-test: wast/wat_arrays_memory/test_array_write_then_read
;; origin: languages/wast/tests/wast/test_wat_arrays_memory.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start") (local $i i32)
          block loop local.get $i i32.const 10 i32.ge_u br_if 1
            local.get $i i32.const 4 i32.mul local.get $i local.get $i i32.mul i32.store
            local.get $i i32.const 1 i32.add local.set $i br 0 end end
          i32.const 28 i32.load call $log))
