;; vybe-test: wast/wat_elem_segments/an_index_spelling_and_an_import_must_resolve_alike
;; vybe-test-mode: run
;;
;; ⛔ A SIBLING SPELLING THAT NOBODY COMPARES IS WHERE THESE BUGS LIVE.
;;
;; Two pairs of bugs shipped because one spelling of a construct resolved and
;; the other silently did not — no error, no wrong answer, just an ABSENT one:
;;
;;   (table funcref (elem 0 1))        pushed the literal "0" as a function
;;                                     NAME, matched nothing, and left the slot
;;                                     null. `(elem $a $b)` worked, so the named
;;                                     form masked it; only func_ptrs.wast in
;;                                     the whole spec suite writes the numeric
;;                                     one.
;;   (table funcref (elem $imported))  qualified the name as `ThisModule.f`,
;;                                     which resolves to nothing for an IMPORT.
;;                                     `call` had always gone through
;;                                     `import_alias`; the inline elem list
;;                                     never did.
;;
;; Both trapped "uninitialized element 0" at the first `call_indirect`, which
;; names the symptom and not the cause. This file pins the AGREEMENT rather
;; than either spelling: every pair below is the same program written two ways,
;; so a regression in one half shows up as a disagreement.

(module
  (type $T (func (result i32)))
  (import "spectest" "print_i32" (func $imported (param i32)))
  (func $a (result i32) (i32.const 11))
  (func $b (result i32) (i32.const 22))

  ;; NUMERIC funcidx in an inline elem — index 1 is $a ($imported is 0).
  (table $numeric funcref (elem 1 2))
  ;; The SAME table written with names.
  (table $named funcref (elem $a $b))
  ;; An inline elem naming an IMPORT.
  (table $withimport funcref (elem $imported $a))

  (func (export "numeric0") (result i32) (call_indirect $numeric (type $T) (i32.const 0)))
  (func (export "named0")   (result i32) (call_indirect $named   (type $T) (i32.const 0)))
  (func (export "numeric1") (result i32) (call_indirect $numeric (type $T) (i32.const 1)))
  (func (export "named1")   (result i32) (call_indirect $named   (type $T) (i32.const 1)))
  ;; The import must occupy its slot rather than leaving it null.
  (func (export "import_slot_filled") (result i32)
    (ref.is_null (table.get $withimport (i32.const 0))))
  (func (export "local_after_import") (result i32)
    (call_indirect $withimport (type $T) (i32.const 1)))
)

;; Numeric and named spellings must give the SAME answer.
(assert_return (invoke "numeric0") (i32.const 11))
(assert_return (invoke "named0")   (i32.const 11))
(assert_return (invoke "numeric1") (i32.const 22))
(assert_return (invoke "named1")   (i32.const 22))
;; An imported function fills its slot like a defined one.
(assert_return (invoke "import_slot_filled") (i32.const 0))
(assert_return (invoke "local_after_import") (i32.const 11))
