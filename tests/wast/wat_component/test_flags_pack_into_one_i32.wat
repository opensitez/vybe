;; vybe-test: wast/wat_component/test_flags_pack_into_one_i32
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §Flat Lifting (:3397)  lift_flat_flags — `i = vi.next('i32')`
;;   §Flat Lowering (:3504) lower_flat_flags — `[pack_flags_into_int(v, labels)]`
;;   §Despecialization (:2185) — `flags` is DELIBERATELY not expanded.
;;
;; ▶▶ `flags` IS NOT A RECORD OF `bool`, AND THE ARITY PROVES IT. It is the one
;; specialized type `despecialize()` refuses to expand, because it BIT-PACKS:
;; one bit per label into a single integer, flattening to exactly ONE core
;; `i32` no matter how many labels there are.
;;
;; A record of three `bool`s would flatten to THREE core `i32`s. The core
;; callee here declares `(param i32)` — one parameter — and the lowered core
;; function does too, so the expansion this test rules out fails on ARITY
;; before it ever produces a wrong number. That is why the callee is
;; single-parameter rather than convenient.
;;
;; 0b101 = 5 is `read` and `exec` set, `write` clear — deliberately not
;; 0b111, so a pack that ignored its input and set every bit would give 7, and
;; not 0b001, so a pack that only ever emitted the first bit would give 1.
;; The callee multiplies by 10, so 50 can only come from 5 surviving the
;; unpack-into-a-record and the re-pack back to an integer.
;;
;; ⛔ WHAT THIS FILE DOES NOT PROVE: bit ORDER. A round trip cannot detect a
;; pack and unpack that are reversed *consistently* — 5 in, record, 5 out,
;; either way. Pinning `bit k == labels[k]` needs the MEMORY path, where a
;; specific byte can be read back; `store_flags` writes only
;; `elem_size_flags(labels)` bytes (1 here, not 4), and nothing yet exercises
;; that. Left explicit rather than implied by a green test.

(component
  (core module $m
    (func (export "scale") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 10))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "scale" (core func $s))

  (type $ft (func (param "f" (flags "read" "write" "exec")) (result u32)))
  (canon lift (core func $s) (func $lifted (type $ft)))
  (canon lower (func $lifted) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 5))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 50))
