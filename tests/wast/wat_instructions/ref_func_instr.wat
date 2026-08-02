;; vybe-test: wast/wat_instructions/ref_func_instr
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func $f) (func ref.func $f drop))
