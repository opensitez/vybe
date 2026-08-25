;; vybe-test: wast/wat_component/test_flags_beyond_thirty_two_labels_refuses
;; vybe-test-mode: run-fail
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md:2294
;;
;;   def alignment_flags(labels):
;;     n = len(labels)
;;     assert(0 < n <= 32)
;;
;; ⛔ `run-fail` is green on ANY failure, so the message MUST be read. It is:
;;
;;   `flags` with 33 labels — the canonical ABI packs them into ONE i32, so the
;;   spec asserts `0 < n <= 32` (CanonicalABI.md:2294). Beyond 32 there is no
;;   bit to carry the flag and it would be dropped silently
;;
;; ▶▶ THE BOUND IS NOT STYLISTIC. `flags` flattens to exactly one core `i32`
;; (`lower_flat_flags` returns a one-element list), so label 33 has no bit to
;; live in. `pack_flags_into_int` would shift it out and `unpack_flags_from_int`
;; would read it back as `false` — a flag the source set, silently clear on the
;; other side, with every layer in between reporting success.
;;
;; That silence is why this refuses at COMPILE time rather than trapping at the
;; boundary: by the time a value crosses, the information is already gone and
;; there is nothing left to report.
;;
;; Exactly 33 labels — one past the limit — so the test cannot pass by the
;; bound being enforced somewhere loosely.

(component
  (type $ft (func (param "f" (flags
    "a1" "a2" "a3" "a4" "a5" "a6" "a7" "a8"
    "b1" "b2" "b3" "b4" "b5" "b6" "b7" "b8"
    "c1" "c2" "c3" "c4" "c5" "c6" "c7" "c8"
    "d1" "d2" "d3" "d4" "d5" "d6" "d7" "d8"
    "e1")) (result u32)))
  (canon lift (core func 999) (func $f (type $ft)))
)
