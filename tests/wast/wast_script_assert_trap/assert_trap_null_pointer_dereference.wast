;; vybe-test: wast/wast_script_assert_trap/assert_trap_null_pointer_dereference
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

(module (type $S (struct (field i32))) (func (export "f") ref.null $S struct.get $S 0 drop))
(assert_trap (invoke "f") "null pointer dereference")
