;; vybe-test: wast/wat_stack_switching/cont_bind_instruction
;; origin: languages/wast/tests/wast/test_wat_stack_switching.rs
;; vybe-test-mode: compile

(module (type $ft (func (param i32))) (type $ct (cont $ft))
          (type $ft2 (func)) (type $ct2 (cont $ft2))
          (func $f (param i32))
          (func (export "_start") ref.func $f cont.new $ct
            i32.const 1 cont.bind $ct $ct2 drop))
