;; vybe-test: wast/wat_global_state/test_global_decrement_loop
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (global $n (mut i32) (i32.const 3))
        (func (export "_start")
          block loop
            global.get $n i32.eqz br_if 1
            global.get $n i32.const 1 i32.sub global.set $n br 0
          end end
          global.get $n i32.const 0 call $vybe_check_i32))
