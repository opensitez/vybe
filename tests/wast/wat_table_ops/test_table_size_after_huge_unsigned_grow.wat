;; vybe-test: wast/wat_table_ops/test_table_size_after_huge_unsigned_grow
;; origin: proposals/spec/test/core/table_grow.wast (spec-compliance regression)
;;
;; A `table.grow` that reports -1 must leave the table's size untouched.

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
  (table $t 16 funcref)
(func (export "_start")
  ref.null func
  i32.const 0xfffffff0
  table.grow $t
  drop
  table.size $t
  i32.const 16 call $vybe_check_i32
)
)
