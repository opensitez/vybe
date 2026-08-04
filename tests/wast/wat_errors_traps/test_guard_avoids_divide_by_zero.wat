;; vybe-test: wast/wat_errors_traps/test_guard_avoids_divide_by_zero
;; origin: languages/wast/tests/wast/test_wat_errors_traps.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (func $safediv (param $a i32) (param $b i32) (result i32)
          local.get $b i32.eqz
          if (result i32) i32.const -1 else local.get $a local.get $b i32.div_s end)
        (func (export "_start") i32.const 10 i32.const 0 call $safediv i32.const -1 call $vybe_check_i32))
