;; vybe-test: wast/wat_more_array_constructors/test_array_len
;; origin: languages/wast/tests/wast/test_wat_more_array_constructors.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start")
          i32.const 0 i32.const 6 array.new $a array.len call $log))
