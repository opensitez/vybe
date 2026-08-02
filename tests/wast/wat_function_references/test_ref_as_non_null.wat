;; vybe-test: wast/wat_function_references/test_ref_as_non_null
;; origin: languages/wast/tests/wast/test_wat_function_references.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (type $t (func (result i32)))
        (func $f (type $t) i32.const 7)
        (func (export "_start") ref.func $f ref.as_non_null call_ref $t call $log))
