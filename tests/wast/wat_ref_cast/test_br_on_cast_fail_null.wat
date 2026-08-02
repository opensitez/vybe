;; vybe-test: wast/wat_ref_cast/test_br_on_cast_fail_null
;; origin: languages/wast/tests/wast/test_wat_ref_cast.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  ref.null $Base
  local.set $s
  
  block (result (ref null $Sub))
    local.get $s
    br_on_cast 0 $Base $Sub
    drop
    i32.const 99
    i32.const 88
    struct.new $Sub
  end
  struct.get $Sub 1
  call $log
)
)
