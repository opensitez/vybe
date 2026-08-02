;; vybe-test: wast/wat_control_flow/throw_ref_instr
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module
  (tag $e)
  (func (export "f")
    (block $h (result exnref)
      (try_table (catch_all_ref $h)
        (nop))
      unreachable)
    throw_ref))
