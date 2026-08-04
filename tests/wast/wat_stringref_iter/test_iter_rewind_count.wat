;; vybe-test: wast/wat_stringref_iter/test_iter_rewind_count
;; origin: languages/wast/tests/wast/test_wat_stringref_iter.rs

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
  (memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start") (local $it (ref null $dummy))
  i32.const 0 i32.const 5 string.new_utf8
  string.as_iter
  local.set $it
  local.get $it i32.const 4 stringview_iter.advance drop  ;; at pos 4
  local.get $it i32.const 2 stringview_iter.rewind        ;; back 2
  i32.const 2 call $vybe_check_i32)
(type $dummy (struct))
)
