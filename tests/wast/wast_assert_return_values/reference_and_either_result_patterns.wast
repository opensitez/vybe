;; vybe-test: wast/wast_assert_return_values/reference_and_either_result_patterns
;; vybe-test-mode: run
;;
;; `assert_return` accepts RESULT PATTERNS, not only values. Four shapes were
;; missing, and every one of them was a PARSE ERROR that took its whole file
;; with it — `gc/array.wast`, `gc/i31.wast`, `gc/extern.wast` and the six
;; `relaxed-simd` files never ran a single assertion.
;;
;;   * `(ref.array)`, `(ref.struct)`, `(ref.i31)`, `(ref.eq)` — "a non-null
;;     reference OF THIS KIND". There is no payload to compare: the check is
;;     `ref.test` against the abstract heap type, in its NON-nullable form, so
;;     a null reference fails it.
;;   * `(ref.host N)` — the payload-carrying spelling `internalize` answers,
;;     the mirror of `(ref.extern N)`.
;;   * `(either r1 r2 …)` — the relaxed-simd proposal's NONDETERMINISM. An
;;     implementation may answer any of the listed results and be conforming,
;;     so the assertion fails only when NONE of them matched.
;;
;; Written so a pattern that silently degraded to "anything goes" would be
;; caught: each kind is also asserted NOT to match a value of another kind.

(module
  (type $vec (array i32))
  (type $pt (struct (field i32)))

  (func (export "an_array") (result anyref) (array.new_default $vec (i32.const 3)))
  (func (export "a_struct") (result anyref) (struct.new_default $pt))
  (func (export "an_i31") (result anyref) (ref.i31 (i32.const 7)))
  (func (export "a_null") (result anyref) (ref.null any))

  ;; The same claims stated through `ref.test`. A wast action takes CONSTANT
  ;; arguments only, so each pairing is its own export rather than a nested
  ;; invoke.
  (func (export "array_is_array") (result i32)
    (ref.test (ref array) (array.new_default $vec (i32.const 3))))
  (func (export "struct_is_array") (result i32)
    (ref.test (ref array) (struct.new_default $pt)))
  (func (export "i31_is_array") (result i32)
    (ref.test (ref array) (ref.i31 (i32.const 7))))
  (func (export "null_is_array") (result i32)
    (ref.test (ref array) (ref.null any)))
  (func (export "struct_is_struct") (result i32)
    (ref.test (ref struct) (struct.new_default $pt)))
  (func (export "array_is_struct") (result i32)
    (ref.test (ref struct) (array.new_default $vec (i32.const 3))))
  (func (export "i31_is_i31") (result i32)
    (ref.test (ref i31) (ref.i31 (i32.const 7))))
  (func (export "array_is_i31") (result i32)
    (ref.test (ref i31) (array.new_default $vec (i32.const 3))))
)

;; The patterns themselves.
(assert_return (invoke "an_array") (ref.array))
(assert_return (invoke "an_array") (ref.eq))
(assert_return (invoke "a_struct") (ref.struct))
(assert_return (invoke "a_struct") (ref.eq))
(assert_return (invoke "an_i31") (ref.i31))
(assert_return (invoke "an_i31") (ref.eq))
(assert_return (invoke "a_null") (ref.null any))

;; …and the same claims stated through `ref.test`, so a pattern that stopped
;; discriminating would disagree with the instruction it is defined as.
(assert_return (invoke "array_is_array") (i32.const 1))
(assert_return (invoke "struct_is_array") (i32.const 0))
(assert_return (invoke "i31_is_array") (i32.const 0))
(assert_return (invoke "null_is_array") (i32.const 0))
(assert_return (invoke "struct_is_struct") (i32.const 1))
(assert_return (invoke "array_is_struct") (i32.const 0))
(assert_return (invoke "i31_is_i31") (i32.const 1))
(assert_return (invoke "array_is_i31") (i32.const 0))

;; ── `(either …)` ────────────────────────────────────────────────────
;; A deterministic function, asserted against a set containing its answer and
;; against a set that does not — the second is what proves `either` is not
;; simply accepting everything.
(module
  (func (export "two") (result i32) (i32.const 2))
  (func (export "vec") (result v128) (v128.const i32x4 1 2 3 4)))

(assert_return (invoke "two") (either (i32.const 1) (i32.const 2) (i32.const 3)))
(assert_return (invoke "two") (either (i32.const 2)))
(assert_return (invoke "vec")
  (either (v128.const i32x4 9 9 9 9) (v128.const i32x4 1 2 3 4)))

;; ── `(ref.host N)` and `(ref.extern N)` ─────────────────────────────
(module
  (func (export "internalize") (param externref) (result anyref)
    (any.convert_extern (local.get 0)))
  (func (export "externalize") (param anyref) (result externref)
    (extern.convert_any (local.get 0))))

(assert_return (invoke "internalize" (ref.extern 1)) (ref.host 1))
(assert_return (invoke "externalize" (ref.host 2)) (ref.extern 2))
