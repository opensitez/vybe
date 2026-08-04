;; vybe-test: wast/wat_i64_relational/test_i64_eq_false
;; origin: languages/wast/tests/wast/test_wat_i64_relational.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
  i64.const 42
  i64.const 43
  i64.eq
  i32.const 0 call $vybe_check_i32
)
)
