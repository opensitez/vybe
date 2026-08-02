;; vybe-test: wast/wat_instructions/compile_factorial_folded
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module
  (func $factorial (export "factorial") (param $n i32) (result i32)
    (if (result i32) (i32.le_s (local.get $n) (i32.const 1))
      (then (i32.const 1))
      (else (i32.mul (local.get $n)
                     (call $factorial (i32.sub (local.get $n) (i32.const 1)))))))
)
