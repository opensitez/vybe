;; vybe-test: wast/wat_function_references/test_call_ref_with_args
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
        (type $bin (func (param i32 i32) (result i32)))
        (func $add (type $bin) local.get 0 local.get 1 i32.add)
        (func (export "_start")
          i32.const 20 i32.const 22 ref.func $add call_ref $bin i32.const 42 call $vybe_check_i32))
