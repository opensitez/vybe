;; vybe-test: wast/wat_function_references/test_function_ref_stored_and_called
;; origin: languages/wast/tests/wast/test_wat_function_references.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (type $t (func (result i32)))
        (func $a (type $t) i32.const 100)
        (func $b (type $t) i32.const 200)
        (func $pick (param $which i32) (result (ref $t))
          local.get $which if (result (ref $t)) ref.func $a else ref.func $b end)
        (func (export "_start") i32.const 0 call $pick call_ref $t call $log))
