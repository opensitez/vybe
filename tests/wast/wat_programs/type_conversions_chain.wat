;; vybe-test: wast/wat_programs/type_conversions_chain
;; origin: languages/wast/tests/wast/test_wat_programs.rs
;; vybe-test-mode: compile

(module
  (func $conv (export "conv") (param $x i32) (result f64)
    local.get $x
    f32.convert_i32_s
    f64.promote_f32)
)
