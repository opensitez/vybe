;; vybe-test: wast/wat_struct_set/test_struct_set
;; origin: languages/wast/tests/wast/test_wat_struct_set.rs

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
  (type $Point (struct (field (mut i32)) (field (mut i32))))
(func (export "_start") (local $p (ref null $Point))
  i32.const 10
  i32.const 20
  struct.new $Point
  local.set $p
  
  local.get $p
  i32.const 42
  struct.set $Point 0
  
  local.get $p
  struct.get $Point 0
  i32.const 42 call $vybe_check_i32
)
)
