;; vybe-test: wast/wat_control_flow/if_compiles
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module
  (func $abs (export "abs") (param $x i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $x) (i32.const 0))
      (then (i32.sub (i32.const 0) (local.get $x)))
      (else (local.get $x)))))
