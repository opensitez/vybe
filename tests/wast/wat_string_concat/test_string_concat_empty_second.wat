;; vybe-test: wast/wat_string_concat/test_string_concat_empty_second
;; origin: languages/wast/tests/wast/test_wat_string_concat.rs

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
  i32.const 0
  string.new_utf8
  
  string.concat
  string.measure_utf8
  i32.const 5 call $vybe_check_i32
)
)
