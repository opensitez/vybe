;; vybe-test: wast/wat_algorithms/test_gcd_recursive
;; origin: languages/wast/tests/wast/test_wat_algorithms.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (func $gcd (param $a i32) (param $b i32) (result i32)
          local.get $b i32.eqz
          if (result i32) local.get $a
          else local.get $b local.get $a local.get $b i32.rem_u call $gcd end)
        (func (export "_start") i32.const 1071 i32.const 462 call $gcd i32.const 21 call $vybe_check_i32))
