;; vybe-test: wast/wat_memory_grow/test_memory_size_after_huge_unsigned_grow
;; origin: proposals/spec/test/core/memory_grow.wast (spec-compliance regression)
;;
;; The other half of the failed-grow contract: a `memory.grow` that reports -1
;; must leave the memory EXACTLY as it was. Checking the return value alone
;; cannot see a grow that half-happened.

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
  drop
  memory.size
  i32.const 1 call $vybe_check_i32
)
)
