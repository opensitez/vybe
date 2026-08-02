;; vybe-test: wast/wat_ref_as_non_null/test_ref_as_non_null_trap_struct
;; origin: languages/wast/tests/wast/test_wat_ref_as_non_null.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $S (struct (field i32)))
(func (export "_start")
  ref.null $S
  ref.as_non_null
  drop
  i32.const 42
  call $log
)
)
