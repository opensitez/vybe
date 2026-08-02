;; vybe-test: wast/wat_array_copy/test_array_copy_oob_dest
;; origin: languages/wast/tests/wast/test_wat_array_copy.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Arr (array (mut i32)))
(func (export "_start") (local $a1 (ref null $Arr)) (local $a2 (ref null $Arr))
  i32.const 10
  i32.const 5
  array.new $Arr
  local.set $a1
  
  i32.const 20
  i32.const 5
  array.new $Arr
  local.set $a2
  
  local.get $a2
  i32.const 4
  local.get $a1
  i32.const 0
  i32.const 3
  array.copy $Arr $Arr
  
  i32.const 0
  call $log
)
)
