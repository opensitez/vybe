;; vybe-test: wast/wat_struct_new/test_struct_new_with_args
;; origin: languages/wast/tests/wast/test_wat_struct_new.rs

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
  (type $Point (struct (field i32) (field i32)))
(func (export "_start") (local $p (ref null $Point))
  i32.const 10
  i32.const 20
  struct.new $Point
  local.set $p
  local.get $p
  struct.get $Point 1
  i32.const 20 call $vybe_check_i32
)
)
