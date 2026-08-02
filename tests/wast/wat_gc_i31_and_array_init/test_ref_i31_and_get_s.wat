;; vybe-test: wast/wat_gc_i31_and_array_init/test_ref_i31_and_get_s
;; origin: languages/wast/tests/wast/test_wat_gc_i31_and_array_init.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") i32.const 42 ref.i31 i31.get_s call $log))
