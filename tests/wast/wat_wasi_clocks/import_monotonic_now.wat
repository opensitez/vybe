;; vybe-test: wast/wat_wasi_clocks/import_monotonic_now
;; origin: languages/wast/tests/wast/test_wat_wasi_clocks.rs
;; vybe-test-mode: compile

(module (import "wasi:clocks/monotonic-clock" "now" (func $now (result i64))))
