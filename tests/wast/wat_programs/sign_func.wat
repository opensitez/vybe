;; vybe-test: wast/wat_programs/sign_func
;; origin: languages/wast/tests/wast/test_wat_programs.rs
;; vybe-test-mode: compile

(module
  (func $sign (export "sign") (param $x i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $x) (i32.const 0))
      (then (i32.const -1))
      (else
        (if (result i32) (i32.gt_s (local.get $x) (i32.const 0))
          (then (i32.const 1))
          (else (i32.const 0))))))
)
