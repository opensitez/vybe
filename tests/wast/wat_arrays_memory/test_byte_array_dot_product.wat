;; vybe-test: wast/wat_arrays_memory/test_byte_array_dot_product
;; origin: languages/wast/tests/wast/test_wat_arrays_memory.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\01\02\03") (data (i32.const 8) "\04\05\06")
        (func (export "_start") (local $i i32) (local $s i32)
          block loop local.get $i i32.const 3 i32.ge_u br_if 1
            local.get $s local.get $i i32.load8_u local.get $i i32.const 8 i32.add i32.load8_u i32.mul i32.add local.set $s
            local.get $i i32.const 1 i32.add local.set $i br 0 end end
          local.get $s call $log))
