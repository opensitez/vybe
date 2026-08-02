;; vybe-test: wast/wat_array_get_set/test_array_get_null
;; origin: languages/wast/tests/wast/test_wat_array_get_set.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Arr (array i32))
(func (export "_start") (local $a (ref null $Arr))
  ref.null $Arr
  local.set $a
  
  local.get $a
  i32.const 0
  array.get $Arr
  call $log
)
)
