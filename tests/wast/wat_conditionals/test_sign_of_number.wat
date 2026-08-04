;; vybe-test: wast/wat_conditionals/test_sign_of_number
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
        (func $sign (param $x i32) (result i32)
          local.get $x i32.const 0 i32.gt_s
          if (result i32) i32.const 1
          else local.get $x i32.const 0 i32.lt_s
               if (result i32) i32.const -1 else i32.const 0 end end)
        (func (export "_start") i32.const -7 call $sign i32.const -1 call $vybe_check_i32))
