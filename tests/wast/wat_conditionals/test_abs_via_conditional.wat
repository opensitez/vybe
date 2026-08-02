;; vybe-test: wast/wat_conditionals/test_abs_via_conditional
;; origin: languages/wast/tests/wast/test_wat_conditionals.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $abs (param $x i32) (result i32)
          local.get $x i32.const 0 i32.lt_s
          if (result i32) i32.const 0 local.get $x i32.sub else local.get $x end)
        (func (export "_start") i32.const -42 call $abs call $log))
