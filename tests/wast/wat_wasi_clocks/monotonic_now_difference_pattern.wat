;; vybe-test: wast/wat_wasi_clocks/monotonic_now_difference_pattern
;; origin: languages/wast/tests/wast/test_wat_wasi_clocks.rs
;; vybe-test-mode: compile

(module
          (import "wasi:clocks/monotonic-clock" "now" (func $now (result i64)))
          (func (export "_start") (result i64) call $now call $now i64.sub))
