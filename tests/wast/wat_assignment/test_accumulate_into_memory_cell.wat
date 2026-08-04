;; vybe-test: wast/wat_assignment/test_accumulate_into_memory_cell
;; origin: languages/wast/tests/wast/test_wat_assignment.rs

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
        (func $add (param $n i32) i32.const 0 i32.const 0 i32.load local.get $n i32.add i32.store)
        (func (export "_start")
          i32.const 3 call $add i32.const 4 call $add i32.const 5 call $add
          i32.const 0 i32.load i32.const 12 call $vybe_check_i32))
