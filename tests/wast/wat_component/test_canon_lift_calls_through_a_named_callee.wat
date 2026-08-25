;; vybe-test: wast/wat_component/test_canon_lift_calls_through_a_named_callee
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §canon lift (:3632), and Explainer.md §10
;;   `canon ::= (canon lift core-prefix(<core:funcsortidx>) <canonopt>*
;;                    bind-id(<externtype>))`
;;
;; ▶▶ THE DISCRIMINATOR IS 1 vs 9, AND IT EXISTS BECAUSE THE TYPES DISAGREE.
;;
;; The core callee returns `(i32.const 9)`. The lifted type says `bool`. So:
;;
;;   * lifted  → 9 is a non-zero i32 → `true` → lowered back to a core `1`
;;   * skipped → the caller sees the callee's raw `9`
;;
;; Nine is deliberately not 1: with `(result u32)` both paths give 9 and the
;; test proves nothing. An earlier version of this file did exactly that — core
;; callee returns 5, assert says 5 — and was GREEN FOR THE WRONG REASON while
;; `canon lift` was lifting `Null` and the caller was reading the callee's
;; value off the operand stack. Two bugs hid behind it:
;;
;;   1. `call_canon_callee` passed `execute_until(depth)` where every other
;;      caller passes `depth + 1`. The callee's RETURN pops its own frame
;;      first, so frames are back to `depth` when the floor check runs and
;;      `depth < depth` is false — the arm fell through, PUSHED the result and
;;      kept interpreting. Lift then lifted `Null`.
;;   2. `exec_canon_lift` pushed nothing. That is right for the spec —
;;      `canon_lift` delivers through `task.return_` and a core caller reaches
;;      a lifted function via `canon lower` — but this VM's canon-import ABI is
;;      one value per call, so pushing none makes the caller read the slot
;;      below. Importing `lift@N` into core wasm stands in for lower ∘ lift, so
;;      the result is lowered back to flat before it is pushed.
;;
;; Everything else on the path had to exist too: the grammar's `bind-id`, the
;; core function index space, `(alias core export …)`,
;; `CanonCallee::CoreExport` (the VM reads `$callee` as a CHUNK index, which
;; only the compiler assigns), and a producer for `VM::canon_functypes`.
;;
;; `peek` is a second, independent check: the callee ran EXACTLY ONCE. A lift
;; that resolved the wrong chunk, or ran it twice, fails there rather than on
;; the value.

(component
  (core module $m
    (global $g (mut i32) (i32.const 0))
    (func (export "run") (result i32)
      (global.set $g (i32.add (global.get $g) (i32.const 1)))
      (i32.const 9))
    (func (export "peek") (result i32) (global.get $g)))
  (core instance $mi (instantiate $m))
  (alias core export $mi "run" (core func $r))

  (type $ft (func (result bool)))
  (canon lift (core func $r) (func $lifted (type $ft)))

  (core module $caller
    (import "canon" "lift@0" (func $l (result i32)))
    (func (export "fire") (result i32)
      (call $l)))
  (core instance (instantiate $caller))
)

(assert_return (invoke "fire") (i32.const 1))
(assert_return (invoke "peek") (i32.const 1))
