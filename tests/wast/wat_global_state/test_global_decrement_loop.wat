;; vybe-test: wast/wat_global_state/test_global_decrement_loop
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $n (mut i32) (i32.const 3))
        (func (export "_start")
          block loop
            global.get $n i32.eqz br_if 1
            global.get $n i32.const 1 i32.sub global.set $n br 0
          end end
          global.get $n call $log))
