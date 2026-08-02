;; vybe-test: wast/wat_programs/f64_distance
;; origin: languages/wast/tests/wast/test_wat_programs.rs
;; vybe-test-mode: compile

(module
  (func $dist (export "dist") (param $x f64) (param $y f64) (result f64)
    local.get $x local.get $x f64.mul
    local.get $y local.get $y f64.mul
    f64.add
    f64.sqrt)
)
