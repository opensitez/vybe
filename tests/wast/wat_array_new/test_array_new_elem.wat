;; vybe-test: wast/wat_array_new/test_array_new_elem
;; origin: languages/wast/tests/wast/test_wat_array_new.rs

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
  (type $Arr (array (mut funcref)))
(func $f)
(elem $e $f)
(func (export "_start") (local $a (ref null $Arr))
  i32.const 0
  i32.const 1
  array.new_elem $Arr $e
  local.set $a

  local.get $a
  array.len
  i32.const 1 call $vybe_check_i32
)
)
