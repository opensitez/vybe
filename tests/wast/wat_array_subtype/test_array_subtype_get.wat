;; vybe-test: wast/wat_array_subtype/test_array_subtype_get
;; origin: languages/wast/tests/wast/test_wat_array_subtype.rs

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
  (type $Base (array (mut i32)))
(type $Sub (array_subtype (mut i32) $Base))
(func (export "_start") (local $a (ref null $Base))
  i32.const 42
  i32.const 5
  array.new $Sub
  local.set $a
  
  local.get $a
  i32.const 2
  array.get $Base
  i32.const 42 call $vybe_check_i32
)
)
