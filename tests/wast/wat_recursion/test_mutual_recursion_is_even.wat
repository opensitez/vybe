;; vybe-test: wast/wat_recursion/test_mutual_recursion_is_even
;; origin: languages/wast/tests/wast/test_wat_recursion.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $even (param $n i32) (result i32)
          local.get $n i32.eqz
          if (result i32) i32.const 1
          else local.get $n i32.const 1 i32.sub call $odd end)
        (func $odd (param $n i32) (result i32)
          local.get $n i32.eqz
          if (result i32) i32.const 0
          else local.get $n i32.const 1 i32.sub call $even end)
        (func (export "_start") i32.const 10 call $even call $log))
