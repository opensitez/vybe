;; vybe-test: wast/wat_wasi_random/import_get_random_bytes
;; origin: languages/wast/tests/wast/test_wat_wasi_random.rs
;; vybe-test-mode: compile

(module (import "wasi:random/random" "get-random-bytes" (func $r (param i64) (result i32))))
