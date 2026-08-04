;; vybe-test: wast/wat_gc_i31_and_array_init/test_i31_get_u_zero_extends
;; origin: languages/wast/tests/wast/test_wat_gc_i31_and_array_init.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (func (export "_start") i32.const -1 ref.i31 i31.get_u i32.const 2147483647 call $vybe_check_i32))
