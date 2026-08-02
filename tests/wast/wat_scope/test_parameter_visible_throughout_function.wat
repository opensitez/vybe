;; vybe-test: wast/wat_scope/test_parameter_visible_throughout_function
;; origin: languages/wast/tests/wast/test_wat_scope.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $f (param $p i32) (result i32)
          block local.get $p i32.const 100 i32.gt_s if unreachable end end
          local.get $p i32.const 2 i32.mul)
        (func (export "_start") i32.const 21 call $f call $log))
