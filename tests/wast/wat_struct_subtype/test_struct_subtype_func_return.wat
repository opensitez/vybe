;; vybe-test: wast/wat_struct_subtype/test_struct_subtype_func_return
;; origin: languages/wast/tests/wast/test_wat_struct_subtype.rs

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
  (type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func $f1 (result (ref null $Base))
  i32.const 42
  i32.const 88
  struct.new $Sub)
(func (export "_start")
  call $f1
  struct.get $Base 0
  i32.const 42 call $vybe_check_i32
)
)
