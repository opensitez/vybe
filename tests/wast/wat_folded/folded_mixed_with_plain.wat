;; vybe-test: wast/wat_folded/folded_mixed_with_plain
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module
  (func (param $a i32) (param $b i32) (result i32)
    (i32.add (local.get $a) (local.get $b))
    local.get $a
    i32.add))
