;; vybe-test: wast/wat_arrays_memory/test_prefix_sum_last
;; origin: languages/wast/tests/wast/test_wat_arrays_memory.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\01\00\00\00\02\00\00\00\03\00\00\00\04\00\00\00")
        (func (export "_start") (local $i i32) (local $run i32)
          block loop local.get $i i32.const 16 i32.ge_u br_if 1
            local.get $run local.get $i i32.load i32.add local.set $run
            local.get $i local.get $run i32.store
            local.get $i i32.const 4 i32.add local.set $i br 0 end end
          i32.const 12 i32.load call $log))
