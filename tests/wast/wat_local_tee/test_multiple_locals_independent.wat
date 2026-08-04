;; vybe-test: wast/wat_local_tee/test_multiple_locals_independent
;; origin: languages/wast/tests/wast/test_wat_local_tee.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (func (export "_start")
          (local $a i32) (local $b i32) (local $c i32)
          i32.const 1 local.set $a i32.const 2 local.set $b i32.const 3 local.set $c
          local.get $a local.get $b i32.add local.get $c i32.add i32.const 6 call $vybe_check_i32))
