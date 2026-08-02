;; vybe-test: wast/wat_errors_traps/test_bounds_check_before_load
;; origin: languages/wast/tests/wast/test_wat_errors_traps.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func $safeload (param $addr i32) (result i32)
          local.get $addr i32.const 65536 i32.ge_u
          if (result i32) i32.const -1 else local.get $addr i32.load end)
        (func (export "_start") i32.const 999999 call $safeload call $log))
