;; vybe-test: wast/wat_wasi_random/import_get_random_u64
;; origin: languages/wast/tests/wast/test_wat_wasi_random.rs
;; vybe-test-mode: compile

(module (import "wasi:random/random" "get-random-u64" (func $r (result i64))))
