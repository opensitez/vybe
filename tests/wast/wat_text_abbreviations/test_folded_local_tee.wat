;; vybe-test: wast/wat_text_abbreviations/test_folded_local_tee
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
        (func (export "_start") (local $x i32)
          (i32.const 15 call $vybe_check_i32 (i32.add (local.tee $x (i32.const 10)) (i32.const 5)))))
