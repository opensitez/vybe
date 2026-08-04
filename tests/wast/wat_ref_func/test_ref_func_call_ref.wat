;; vybe-test: wast/wat_ref_func/test_ref_func_call_ref
;; origin: languages/wast/tests/wast/test_wat_ref_func.rs

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
  (type $sig (func (result i32)))
(func $f1 (result i32) i32.const 42)
(func (export "_start") (local $r (ref null $sig))
  ref.func $f1
  local.set $r
  
  local.get $r
  call_ref $sig
  i32.const 42 call $vybe_check_i32
)
)
