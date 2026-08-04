;; vybe-test: wast/wat_select/test_select_ref_true
;; origin: languages/wast/tests/wast/test_wat_select.rs

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
  (func $f1)
(func $f2)
(func (export "_start")
  ref.func $f1
  ref.func $f2
  i32.const 1
  select (result funcref)
  ref.is_null
  i32.const 0 call $vybe_check_i32
)
)
