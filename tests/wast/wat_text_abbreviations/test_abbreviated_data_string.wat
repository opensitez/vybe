;; vybe-test: wast/wat_text_abbreviations/test_abbreviated_data_string
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
        (memory 1) (data (i32.const 0) "\2a\00\00\00")
        (func (export "_start") (i32.const 42 call $vybe_check_i32 (i32.load (i32.const 0)))))
