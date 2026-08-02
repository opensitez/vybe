;; vybe-test: wast/wat_gc_i31_and_array_init/test_array_fill
;; origin: languages/wast/tests/wast/test_wat_gc_i31_and_array_init.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start") (local $arr (ref $a))
          i32.const 0 i32.const 5 array.new $a local.set $arr
          local.get $arr i32.const 1 i32.const 9 i32.const 3 array.fill $a
          local.get $arr i32.const 2 array.get $a call $log))
