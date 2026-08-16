;; vybe-test: wast/wat_call_direct/test_call_numeric_funcidx_skips_non_func_imports
;; origin: proposals/spec/test/core/names.wast (spec-compliance regression)
;;
;; Each import kind has its OWN index space (spec §2.5.1). A `(global)`,
;; `(memory)`, `(table)` or `(tag)` import does not consume a funcidx, so it
;; must not shift the functions declared after it.
;;
;; Counting every `(import …)` as a function pushed each real function one
;; slot along per non-func import ahead of it. Nothing caught that while
;; `call <n>` was unimplemented and the index table was only consulted by
;; `(export "e" (func N))` — an off-by-one there needs a module that both
;; exports by index and imports something that isn't a function.
;;
;; Indices here: log imports 0-3, `$vybe_check_i32` 4, `$seven` 5. The global
;; and memory imports in between contribute nothing.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (import "env" "some_global" (global i32))
  (import "env" "some_memory" (memory 1))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func $seven (result i32)
    i32.const 7)
(func (export "_start")
  call 5
  i32.const 7 call $vybe_check_i32
)
)
