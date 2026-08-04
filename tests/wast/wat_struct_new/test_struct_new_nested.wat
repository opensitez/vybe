;; vybe-test: wast/wat_struct_new/test_struct_new_nested
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
(type $Rect (struct (field (ref $Point)) (field (ref $Point))))
(func (export "_start") (local $r (ref null $Rect))
  i32.const 10
  i32.const 20
  struct.new $Point
  
  i32.const 30
  i32.const 40
  struct.new $Point
  
  struct.new $Rect
  local.set $r
  
  local.get $r
  struct.get $Rect 1
  struct.get $Point 0
  i32.const 30 call $vybe_check_i32
)
)
