;; vybe-test: wast/wat_ref_null/test_ref_is_null_false
;; origin: languages/wast/tests/wast/test_wat_ref_null.rs

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
  (type $t (struct (field i32)))
(func (export "_start")
  (local $s (ref null $t))
  (local.set $s (struct.new $t (i32.const 42)))
  (ref.is_null (local.get $s))
  i32.const 0 call $vybe_check_i32
)
)
