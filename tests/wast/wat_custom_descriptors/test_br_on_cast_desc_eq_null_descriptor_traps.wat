;; vybe-test: wast/wat_custom_descriptors/test_br_on_cast_desc_eq_null_descriptor_traps
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/br_on_cast_desc_eq.wast
;; vybe-test-mode: run-fail

;; `(assert_trap (invoke "self-nullable-val-null") "null descriptor reference")`.
;;
;; The descriptor is checked BEFORE the reference is looked at: the companion
;; assertion `self-nullable-null-null` passes a null reference AND a null
;; descriptor and is asserted to trap rather than to branch, so a null
;; descriptor can never be treated as "just doesn't match".
;;
;; ⚠ WHAT THIS TEST DOES AND DOES NOT PIN. It pins THAT the trap happens. It
;; does NOT pin the message: `br_on_cast_desc_eq` is lowered structurally (the
;; same as `br_on_cast`, whose branch has to integrate with this walker's
;; block/result-temp discipline), so the trap it raises is `unreachable`, and
;; the message reads "unreachable executed" rather than the proposal's "null
;; descriptor reference". The message is only reachable from a VM opcode, and
;; nothing emits `Op::BR_ON_CAST_DESC_EQ` — which DOES carry the right wording.
;; `run-fail` would not have caught the difference either way: it is green on
;; any failure, and `walk_assert_trap` parses the expected message and drops
;; it. Both halves of that are pre-existing gaps, not ones this test creates.

(module
  (rec
    (type $a (descriptor $a.desc) (struct (field i32)))
    (type $a.desc (describes $a) (struct (field i32)))
  )

  (func (export "_start") (local $d (ref null $a.desc)) (local $o (ref null $a))
    i32.const 7
    struct.new $a.desc
    local.set $d

    i32.const 1
    local.get $d
    struct.new_desc $a
    local.set $o

    (block (result anyref)
      (br_on_cast_desc_eq 0 anyref (ref null $a) (local.get $o) (ref.null none))
      (return)
    )
    drop
  )
)
