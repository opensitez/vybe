;; vybe-test: wast/wat_call_direct/test_call_numeric_funcidx_recursive
;; origin: proposals/spec/test/core/fac.wast (spec-compliance regression)
;;
;; A function calling ITSELF by its own numeric index — the shape `fac.wast`
;; opens with. Distinct from the plain numeric-call case: the index has to
;; resolve to a function that is still being walked, so a resolver that only
;; knows about already-finished definitions answers nothing here.

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
  (func $fact (param $n i32) (result i32)
    (if (result i32) (i32.le_s (local.get $n) (i32.const 1))
      (then (i32.const 1))
      (else
        (i32.mul
          (local.get $n)
          (call 5 (i32.sub (local.get $n) (i32.const 1)))
        )
      )
    ))
(func (export "_start")
  i32.const 5
  call $fact
  i32.const 120 call $vybe_check_i32
)
)
