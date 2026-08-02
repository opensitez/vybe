;; vybe-test: wast/wat_stack_switching/resume_instruction
;; origin: languages/wast/tests/wast/test_wat_stack_switching.rs
;; vybe-test-mode: compile

(module (type $ft (func)) (type $ct (cont $ft)) (tag $yield)
          (func $f)
          (func (export "_start") ref.func $f cont.new $ct resume $ct (on $yield $h))
          (func $h))
