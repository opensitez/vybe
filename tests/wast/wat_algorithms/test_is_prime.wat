;; vybe-test: wast/wat_algorithms/test_is_prime
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
        (func $prime (param $n i32) (result i32) (local $i i32)
          local.get $n i32.const 2 i32.lt_s if i32.const 0 return end
          i32.const 2 local.set $i
          block loop
            local.get $i local.get $i i32.mul local.get $n i32.gt_s br_if 1
            local.get $n local.get $i i32.rem_u i32.eqz if i32.const 0 return end
            local.get $i i32.const 1 i32.add local.set $i br 0
          end end i32.const 1)
        (func (export "_start") i32.const 97 call $prime i32.const 1 call $vybe_check_i32))
