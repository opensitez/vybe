;; vybe-test: wast/wat_function_references/test_call_ref_invokes_function
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
        (func $answer (type $t) i32.const 42)
        (func (export "_start") ref.func $answer call_ref $t i32.const 42 call $vybe_check_i32))
