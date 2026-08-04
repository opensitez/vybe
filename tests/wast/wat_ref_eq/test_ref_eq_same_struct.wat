;; vybe-test: wast/wat_ref_eq/test_ref_eq_same_struct
;; origin: languages/wast/tests/wast/test_wat_ref_eq.rs

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
  i32.const 42
  struct.new $S
  local.tee $s
  local.get $s
  ref.eq
  i32.const 1 call $vybe_check_i32
)
)
