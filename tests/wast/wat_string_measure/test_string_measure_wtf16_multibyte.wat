;; vybe-test: wast/wat_string_measure/test_string_measure_wtf16_multibyte
;; origin: languages/wast/tests/wast/test_wat_string_measure.rs

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
(data (i32.const 0) "\e2\82\ac") ;; euro sign, 3 bytes in utf8, 1 code unit in wtf16
(func (export "_start")
  i32.const 0
  i32.const 3
  string.new_utf8
  string.measure_wtf16
  i32.const 1 call $vybe_check_i32
)
)
