;; vybe-test: wast/wat_more_array_constructors/test_array_new_fixed
;; origin: languages/wast/tests/wast/test_wat_more_array_constructors.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start")
          i32.const 10 i32.const 20 i32.const 30 array.new_fixed $a 3
          i32.const 1 array.get $a call $log))
