;; vybe-test: wast/wat_struct_get/test_default_local_is_null
;; origin: languages/wast/tests/wast/test_wat_struct_get.rs

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
  (type $Point (struct (field i32) (field i32)))
(func (export "_start") (local $p (ref null $Point))
  (ref.is_null (local.get $p))
  i32.const 1 call $vybe_check_i32
)
)
