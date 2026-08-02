;; vybe-test: wast/wat_local_tee/test_param_and_local_mix
;; origin: languages/wast/tests/wast/test_wat_local_tee.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $f (param $p i32) (result i32)
          (local $tmp i32) local.get $p i32.const 100 i32.add local.set $tmp local.get $tmp)
        (func (export "_start") i32.const 5 call $f call $log))
