;; vybe-test: wast/wat_text_abbreviations/test_folded_local_tee
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") (local $x i32)
          (call $log (i32.add (local.tee $x (i32.const 10)) (i32.const 5)))))
