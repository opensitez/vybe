;; vybe-test: wast/wat_wasi_clocks/import_monotonic_get_resolution
;; origin: languages/wast/tests/wast/test_wat_wasi_clocks.rs
;; vybe-test-mode: compile

(module (import "wasi:clocks/monotonic-clock" "get-resolution" (func $res (result i64))))
