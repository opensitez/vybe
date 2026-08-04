;; vybe-test: wast/wat_text_abbreviations/test_folded_nested_arithmetic
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (func (export "_start")
          (i32.const 20 call $vybe_check_i32 (i32.mul (i32.add (i32.const 2) (i32.const 3)) (i32.const 4)))))
