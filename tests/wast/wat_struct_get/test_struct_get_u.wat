;; vybe-test: wast/wat_struct_get/test_struct_get_u
;; origin: languages/wast/tests/wast/test_wat_struct_get.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $S (struct (field i8) (field i16)))
(func (export "_start") (local $s (ref null $S))
  i32.const 255 ;; 255 as u8
  i32.const 65535 ;; 65535 as u16
  struct.new $S
  local.set $s
  
  local.get $s
  struct.get_u $S 0
  call $log
)
)
