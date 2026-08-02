;; vybe-test: wast/wat_control_flow/try_table_folded
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module
  (tag $e (param i32))
  (func (export "f")
    (block $h
      (try_table (catch_all $h)
        (nop)))))
