;; vybe-test: wast/wat_stack_switching/cont_new_instruction
;; origin: languages/wast/tests/wast/test_wat_stack_switching.rs
;; vybe-test-mode: compile

(module (type $ft (func)) (type $ct (cont $ft))
          (func $f)
          (func (export "_start") ref.func $f cont.new $ct drop))
