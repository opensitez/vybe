;; vybe-test: wast/wat_programs/clamp_func
;; origin: languages/wast/tests/wast/test_wat_programs.rs
;; vybe-test-mode: compile

(module
  (func $clamp (export "clamp") (param $x i32) (param $lo i32) (param $hi i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $x) (local.get $lo))
      (then (local.get $lo))
      (else
        (if (result i32) (i32.gt_s (local.get $x) (local.get $hi))
          (then (local.get $hi))
          (else (local.get $x))))))
)
