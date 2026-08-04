;; vybe-test: wast/wat_ref_is_null/test_ref_is_null_after_local_set
;; origin: languages/wast/tests/wast/test_wat_ref_is_null.rs

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
(func (export "_start") (local $s (ref null $S))
  ref.null $S
  local.set $s
  local.get $s
  ref.is_null
  i32.const 1 call $vybe_check_i32
)
)
