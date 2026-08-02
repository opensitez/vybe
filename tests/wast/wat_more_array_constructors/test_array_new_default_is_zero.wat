;; vybe-test: wast/wat_more_array_constructors/test_array_new_default_is_zero
;; origin: languages/wast/tests/wast/test_wat_more_array_constructors.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start")
          i32.const 4 array.new_default $a i32.const 0 array.get $a call $log))
