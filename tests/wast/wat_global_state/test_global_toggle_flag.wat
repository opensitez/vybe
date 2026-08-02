;; vybe-test: wast/wat_global_state/test_global_toggle_flag
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $flag (mut i32) (i32.const 0))
        (func $toggle global.get $flag i32.eqz global.set $flag)
        (func (export "_start") call $toggle call $toggle call $toggle global.get $flag call $log))
