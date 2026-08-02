;; vybe-test: wast/wat_struct_get/test_struct_get_s_16
;; origin: languages/wast/tests/wast/test_wat_struct_get.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $S (struct (field i8) (field i16)))
(func (export "_start") (local $s (ref null $S))
  i32.const 255 ;; -1 as i8
  i32.const 65535 ;; -1 as i16
  struct.new $S
  local.set $s
  
  local.get $s
  struct.get_s $S 1
  call $log
)
)
