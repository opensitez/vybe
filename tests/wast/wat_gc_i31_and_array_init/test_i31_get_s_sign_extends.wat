;; vybe-test: wast/wat_gc_i31_and_array_init/test_i31_get_s_sign_extends
;; origin: languages/wast/tests/wast/test_wat_gc_i31_and_array_init.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") i32.const -1 ref.i31 i31.get_s call $log))
