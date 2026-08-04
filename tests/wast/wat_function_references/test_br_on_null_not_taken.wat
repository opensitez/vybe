;; vybe-test: wast/wat_function_references/test_br_on_null_not_taken
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
        (func $f (type $t) i32.const 5)
        (func (export "_start") (result i32)
          block (result (ref $t))
            ref.func $f br_on_null 0
            call_ref $t return
          end drop i32.const -1)
        (func (export "_run") i32.const 5 call $vybe_check_i32))
