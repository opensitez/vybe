;; vybe-test: wast/wat_array_subtype/test_array_subtype_func_param
;; origin: languages/wast/tests/wast/test_wat_array_subtype.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Base (array (mut i32)))
(type $Sub (array_subtype (mut i32) $Base))
(func $f1 (param $a (ref null $Base)) (result i32)
  local.get $a
  i32.const 0
  array.get $Base)
(func (export "_start")
  i32.const 99
  i32.const 5
  array.new $Sub
  call $f1
  call $log
)
)
