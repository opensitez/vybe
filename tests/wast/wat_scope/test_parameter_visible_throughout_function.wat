;; vybe-test: wast/wat_scope/test_parameter_visible_throughout_function
;; origin: languages/wast/tests/wast/test_wat_scope.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (func $f (param $p i32) (result i32)
          block local.get $p i32.const 100 i32.gt_s if unreachable end end
          local.get $p i32.const 2 i32.mul)
        (func (export "_start") i32.const 21 call $f i32.const 42 call $vybe_check_i32))
