;; vybe-test: wast/wat_folded/compile_folded_min
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module
  (func $min (export "min") (param $a i32) (param $b i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $a) (local.get $b))
      (then (local.get $a))
      (else (local.get $b)))))
