;; vybe-test: wast/wat_func_locals/test_local_tee_and_get
;; origin: languages/wast/tests/wast/test_wat_func_locals.rs

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
  (func (export "_start") (local $x i32)
  i32.const 42
  local.tee $x
  local.get $x
  i32.add
  i32.const 84 call $vybe_check_i32
)
)
