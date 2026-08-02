;; vybe-test: wast/wat_execution_extended/test_call_indirect_signature_check
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs
;; vybe-test-mode: compile

(module
  (type $t_void (func))
  (type $t_i32 (func (result i32)))
  (table 1 funcref)
  (func $f (result i32) i32.const 123)
  (elem (i32.const 0) $f)
  (func (export "run")
    i32.const 0
    call_indirect (type $t_void)))
(assert_trap (invoke "run") "indirect call signature mismatch")
