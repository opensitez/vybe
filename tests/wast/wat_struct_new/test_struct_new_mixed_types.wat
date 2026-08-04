;; vybe-test: wast/wat_struct_new/test_struct_new_mixed_types
;; origin: languages/wast/tests/wast/test_wat_struct_new.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
  (type $Mixed (struct (field i32) (field f32) (field i64) (field f64)))
(func (export "_start") (local $m (ref null $Mixed))
  i32.const 42
  f32.const 3.14
  i64.const 99
  f64.const 2.71
  struct.new $Mixed
  local.set $m
  
  local.get $m
  struct.get $Mixed 2
  i64.const 99 call $vybe_check_i64
)
)
