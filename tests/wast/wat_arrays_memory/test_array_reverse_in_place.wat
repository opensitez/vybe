;; vybe-test: wast/wat_arrays_memory/test_array_reverse_in_place
;; origin: languages/wast/tests/wast/test_wat_arrays_memory.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\01\00\00\00\02\00\00\00\03\00\00\00\04\00\00\00")
        (func (export "_start") (local $l i32) (local $r i32) (local $t i32)
          i32.const 12 local.set $r
          block loop local.get $l local.get $r i32.ge_u br_if 1
            local.get $l i32.load local.set $t
            local.get $l local.get $r i32.load i32.store
            local.get $r local.get $t i32.store
            local.get $l i32.const 4 i32.add local.set $l
            local.get $r i32.const 4 i32.sub local.set $r br 0 end end
          i32.const 0 i32.load call $log))
