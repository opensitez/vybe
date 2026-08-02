;; vybe-test: wast/wat_array_get_set/test_array_get_s
;; origin: languages/wast/tests/wast/test_wat_array_get_set.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Arr (array i8))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 255 ;; -1 as i8
  i32.const 1
  array.new $Arr
  local.set $a
  
  local.get $a
  i32.const 0
  array.get_s $Arr
  call $log
)
)
