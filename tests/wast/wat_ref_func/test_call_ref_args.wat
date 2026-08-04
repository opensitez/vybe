;; vybe-test: wast/wat_ref_func/test_call_ref_args
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
  (type $sig (func (param i32 i32) (result i32)))
(func $add (param i32 i32) (result i32)
  local.get 0
  local.get 1
  i32.add)
(func (export "_start") (local $r (ref null $sig))
  ref.func $add
  local.set $r
  
  i32.const 10
  i32.const 20
  local.get $r
  call_ref $sig
  i32.const 30 call $vybe_check_i32
)
)
