;; vybe-test: wast/wat_multivalue/test_func_multi_return_used_in_arithmetic
;; origin: languages/wast/tests/wast/test_wat_multivalue.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $divmod (param i32 i32) (result i32 i32)
          local.get 0 local.get 1 i32.div_u
          local.get 0 local.get 1 i32.rem_u)
        (func (export "_start")
          i32.const 17 i32.const 5 call $divmod i32.add call $log))
