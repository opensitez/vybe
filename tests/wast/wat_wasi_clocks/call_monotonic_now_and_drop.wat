;; vybe-test: wast/wat_wasi_clocks/call_monotonic_now_and_drop
;; origin: languages/wast/tests/wast/test_wat_wasi_clocks.rs
;; vybe-test-mode: compile

(module
          (import "wasi:clocks/monotonic-clock" "now" (func $now (result i64)))
          (func (export "_start") call $now drop))
