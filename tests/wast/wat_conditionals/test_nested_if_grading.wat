;; vybe-test: wast/wat_conditionals/test_nested_if_grading
;; origin: languages/wast/tests/wast/test_wat_conditionals.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (func $grade (param $s i32) (result i32)
          local.get $s i32.const 90 i32.ge_s
          if (result i32) i32.const 4
          else local.get $s i32.const 80 i32.ge_s
               if (result i32) i32.const 3
               else local.get $s i32.const 70 i32.ge_s
                    if (result i32) i32.const 2 else i32.const 1 end end end)
        (func (export "_start") i32.const 85 call $grade i32.const 3 call $vybe_check_i32))
