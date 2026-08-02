;; vybe-test: wast/wast_script_assert_trap/assert_trap_cast_error
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

(module 
  (type $Base (struct (field i32)))
  (type $Sub (struct_subtype (field i32) (field i32) $Base))
  (func (export "f") 
    i32.const 0 struct.new $Base
    ref.cast $Sub drop)
)
(assert_trap (invoke "f") "cast error")
