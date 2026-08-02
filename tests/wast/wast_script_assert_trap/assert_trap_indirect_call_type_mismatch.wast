;; vybe-test: wast/wast_script_assert_trap/assert_trap_indirect_call_type_mismatch
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

(module 
  (type $t1 (func (result i32)))
  (type $t2 (func (result f32)))
  (table 1 funcref)
  (func $g (type $t1) i32.const 0)
  (elem (i32.const 0) $g)
  (func (export "f") i32.const 0 call_indirect (type $t2) drop)
)
(assert_trap (invoke "f") "indirect call type mismatch")
