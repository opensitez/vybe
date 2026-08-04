;; vybe-test: wast/wat_table_ops/test_table_size_after_grow
;; origin: languages/wast/tests/wast/test_wat_table_ops.rs

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
  (table 5 funcref)
(func (export "_start")
  ref.null func
  i32.const 3
  table.grow 0
  drop
  table.size 0
  i32.const 8 call $vybe_check_i32
)
)
