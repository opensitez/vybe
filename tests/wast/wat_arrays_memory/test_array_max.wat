;; vybe-test: wast/wat_arrays_memory/test_array_max
;; origin: languages/wast/tests/wast/test_wat_arrays_memory.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\03\00\00\00\09\00\00\00\05\00\00\00\07\00\00\00")
        (func (export "_start") (local $i i32) (local $m i32)
          block loop local.get $i i32.const 16 i32.ge_u br_if 1
            local.get $i i32.load local.get $m i32.gt_s
            if local.get $i i32.load local.set $m end
            local.get $i i32.const 4 i32.add local.set $i br 0 end end
          local.get $m call $log))
