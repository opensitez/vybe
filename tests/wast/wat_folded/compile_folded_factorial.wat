;; vybe-test: wast/wat_folded/compile_folded_factorial
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module
  (func $fact (export "fact") (param $n i32) (result i32)
    (if (result i32) (i32.le_s (local.get $n) (i32.const 1))
      (then (i32.const 1))
      (else (i32.mul (local.get $n)
                     (call $fact (i32.sub (local.get $n) (i32.const 1))))))))
