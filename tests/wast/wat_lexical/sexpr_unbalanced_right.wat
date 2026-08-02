;; vybe-test: wast/wat_lexical/sexpr_unbalanced_right
;; origin: languages/wast/tests/wast/test_wat_lexical.rs
;; vybe-test-mode: compile-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (
)
