;; vybe-test: wast/wat_struct_get/test_struct_get_u_16
;; origin: languages/wast/tests/wast/test_wat_struct_get.rs

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
  (type $S (struct (field i8) (field i16)))
(func (export "_start") (local $s (ref null $S))
  i32.const 255 ;; 255 as u8
  i32.const 65535 ;; 65535 as u16
  struct.new $S
  local.set $s
  
  local.get $s
  struct.get_u $S 1
  i32.const 65535 call $vybe_check_i32
)
)
