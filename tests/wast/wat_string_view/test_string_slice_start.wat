;; vybe-test: wast/wat_string_view/test_string_slice_start
;; origin: languages/wast/tests/wast/test_wat_string_view.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "hello world")
(func (export "_start")
  i32.const 0
  i32.const 11
  string.new_utf8        ;; "hello world"
  string.as_wtf8
  i32.const 0
  i32.const 5
  stringview_wtf8.slice  ;; "hello"
  string.measure_utf8    ;; length 5
  call $log
)
)
