;; vybe-test: wast/wat_more_array_constructors/test_array_set_then_get
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
        (func (export "_start") (local $arr (ref $a))
          i32.const 0 i32.const 5 array.new $a local.set $arr
          local.get $arr i32.const 2 i32.const 42 array.set $a
          local.get $arr i32.const 2 array.get $a i32.const 42 call $vybe_check_i32))
