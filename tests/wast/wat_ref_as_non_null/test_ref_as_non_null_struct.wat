;; vybe-test: wast/wat_ref_as_non_null/test_ref_as_non_null_struct
;; origin: languages/wast/tests/wast/test_wat_ref_as_non_null.rs

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
  (type $S (struct (field i32)))
(func (export "_start")
  i32.const 42
  struct.new $S
  ref.as_non_null
  struct.get $S 0
  i32.const 42 call $vybe_check_i32
)
)
