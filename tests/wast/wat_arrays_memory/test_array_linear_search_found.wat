;; vybe-test: wast/wat_arrays_memory/test_array_linear_search_found
;; origin: languages/wast/tests/wast/test_wat_arrays_memory.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\0a\00\00\00\14\00\00\00\1e\00\00\00\28\00\00\00")
        (func (export "_start") (local $i i32)
          block loop local.get $i i32.const 16 i32.ge_u
            if i32.const -1 call $log return end
            local.get $i i32.load i32.const 30 i32.eq
            if local.get $i i32.const 4 i32.div_u call $log return end
            local.get $i i32.const 4 i32.add local.set $i br 0 end end))
