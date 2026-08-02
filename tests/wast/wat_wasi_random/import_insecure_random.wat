;; vybe-test: wast/wat_wasi_random/import_insecure_random
;; origin: languages/wast/tests/wast/test_wat_wasi_random.rs
;; vybe-test-mode: compile

(module (import "wasi:random/insecure" "get-insecure-random-u64" (func $r (result i64))))
