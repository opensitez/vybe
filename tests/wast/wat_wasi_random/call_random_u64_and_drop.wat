;; vybe-test: wast/wat_wasi_random/call_random_u64_and_drop
;; origin: languages/wast/tests/wast/test_wat_wasi_random.rs
;; vybe-test-mode: compile

(module
          (import "wasi:random/random" "get-random-u64" (func $r (result i64)))
          (func (export "_start") call $r drop))
