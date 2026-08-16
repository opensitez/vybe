;; vybe-test: wast/wat_memory_grow/test_memory_grow_huge_unsigned_pages
;; origin: proposals/spec/test/core/memory_grow.wast (spec-compliance regression)
;;
;; The page-count operand of `memory.grow` is read UNSIGNED: 0xffffffff is
;; 4294967295 pages, not -1 and not 0. An impossible delta REPORTS failure —
;; it returns -1 and leaves the memory alone — it does not trap.
;;
;; Clamping the operand with `.max(0)` turned this into `memory.grow 0`, which
;; succeeds and returns the current size. The assertion below is the one that
;; separates the two: 0 (the old answer) is a perfectly plausible size.

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
(func (export "_start")
  i32.const 0xffffffff
  memory.grow
  i32.const -1 call $vybe_check_i32
)
)
