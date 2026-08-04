;; vybe-test: wast/wat_function_references/test_ref_as_non_null
;; origin: languages/wast/tests/wast/test_wat_function_references.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (type $t (func (result i32)))
        (func $f (type $t) i32.const 7)
        (func (export "_start") ref.func $f ref.as_non_null call_ref $t i32.const 7 call $vybe_check_i32))
