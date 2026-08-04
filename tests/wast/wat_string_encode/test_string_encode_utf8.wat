;; vybe-test: wast/wat_string_encode/test_string_encode_utf8
;; origin: languages/wast/tests/wast/test_wat_string_encode.rs

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
  (memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  i32.const 10
  string.encode_utf8
  drop
  
  i32.const 10
  i32.load8_u
  i32.const 104 call $vybe_check_i32
)
)
