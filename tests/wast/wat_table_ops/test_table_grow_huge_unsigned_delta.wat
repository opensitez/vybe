;; vybe-test: wast/wat_table_ops/test_table_grow_huge_unsigned_delta
;; origin: proposals/spec/test/core/table_grow.wast (spec-compliance regression)
;;
;; `table.grow`'s delta is UNSIGNED, exactly like `memory.grow`'s page count:
;; 0xfffffff0 is 4294967280 entries. The table has no maximum here, so nothing
;; but the operand's magnitude can make this fail — which is the point. Read as
;; a signed -16 and clamped to 0 it becomes a no-op grow returning the current
;; size (16), and the spec's own `table_grow.wast` asserts -1 for this case.

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
  i32.const -1 call $vybe_check_i32
)
)
