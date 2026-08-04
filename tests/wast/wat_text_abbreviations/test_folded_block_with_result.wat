;; vybe-test: wast/wat_text_abbreviations/test_folded_block_with_result
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
          (i32.const 100 call $vybe_check_i32 (block (result i32) (i32.const 100)))))
