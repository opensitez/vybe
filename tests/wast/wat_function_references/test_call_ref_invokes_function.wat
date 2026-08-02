;; vybe-test: wast/wat_function_references/test_call_ref_invokes_function
;; origin: languages/wast/tests/wast/test_wat_function_references.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (type $t (func (result i32)))
        (func $answer (type $t) i32.const 42)
        (func (export "_start") ref.func $answer call_ref $t call $log))
