;; vybe-test: wast/wat_scope/test_global_shared_across_calls
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
        (global $g (mut i32) (i32.const 0))
        (func $bump global.get $g i32.const 1 i32.add global.set $g)
        (func (export "_start") call $bump call $bump call $bump global.get $g i32.const 3 call $vybe_check_i32))
