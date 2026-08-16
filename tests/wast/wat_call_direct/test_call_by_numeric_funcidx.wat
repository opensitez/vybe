;; vybe-test: wast/wat_call_direct/test_call_by_numeric_funcidx
;; origin: proposals/spec/test/core/fac.wast (spec-compliance regression)
;;
;; `call` takes a funcidx, and a funcidx is a NUMBER — `$name` is the
;; abbreviation, not the other way round. The spec's own suite writes it plainly
;; (`fac.wast`: `(call 0 …)` recursing into itself), and every such call used to
;; be lowered with the integer literal itself as the callee, dying with
;; "f64 is not callable".
;;
;; Index space per spec §2.5.1: IMPORTS FIRST, then defined functions. The four
;; log imports occupy 0-3, so `$vybe_check_i32` is 4 and `$seven` is 5 — which
;; makes `call 5` also the assertion that imports are counted.

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
  (func $seven (result i32)
    i32.const 7)
(func (export "_start")
  call 5
  i32.const 7 call $vybe_check_i32
)
)
