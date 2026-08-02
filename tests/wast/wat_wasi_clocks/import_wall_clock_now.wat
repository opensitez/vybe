;; vybe-test: wast/wat_wasi_clocks/import_wall_clock_now
;; origin: languages/wast/tests/wast/test_wat_wasi_clocks.rs
;; vybe-test-mode: compile

(module (import "wasi:clocks/wall-clock" "now" (func $now (result i64 i32))))
