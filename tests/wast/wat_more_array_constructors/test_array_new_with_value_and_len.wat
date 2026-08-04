;; vybe-test: wast/wat_more_array_constructors/test_array_new_with_value_and_len
;; origin: languages/wast/tests/wast/test_wat_more_array_constructors.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (type $a (array (mut i32)))
        (func (export "_start")
          i32.const 7 i32.const 5 array.new $a i32.const 3 array.get $a i32.const 7 call $vybe_check_i32))
