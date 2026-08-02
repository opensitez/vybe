;; vybe-test: wast/wat_arrays_memory/test_array_count_matching
;; origin: languages/wast/tests/wast/test_wat_arrays_memory.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\02\00\00\00\04\00\00\00\05\00\00\00\06\00\00\00")
        (func (export "_start") (local $i i32) (local $c i32)
          block loop local.get $i i32.const 16 i32.ge_u br_if 1
            local.get $i i32.load i32.const 1 i32.and i32.eqz
            if local.get $c i32.const 1 i32.add local.set $c end
            local.get $i i32.const 4 i32.add local.set $i br 0 end end
          local.get $c call $log))
