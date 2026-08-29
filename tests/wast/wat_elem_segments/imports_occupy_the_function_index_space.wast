;; vybe-test: wast/wat_elem_segments/imports_occupy_the_function_index_space
;; vybe-test-mode: run
;;
;; WASM 3.0 §6.4: imports occupy the LOW end of every index space. A module
;; that imports three functions and then defines two numbers them 0,1,2 and
;; 3,4 — so `func 3` in an element list, `ref.func 1`, `call 0` and
;; `(export "e" (func 2))` all mean something an implementation that counts
;; only DEFINED functions cannot name.
;;
;; Three separate holes met in `bulk-memory/table_init.wast`, which imports
;; five functions and then writes its element lists entirely in numbers:
;;
;;   1. An UNNAMED standalone `(import "a" "f" (func …))` took a function index
;;      but nothing ever bound that index to the exporter's method. A named
;;      import worked (its `$id` is bound), so the gap was invisible in every
;;      test that named its imports.
;;   2. An element item written as a NUMBER stayed the literal string — the
;;      active path built a member access on a class member called "3", which
;;      no class has, and the table slot silently kept its null.
;;   3. A PASSIVE segment resolved every item under ONE class, so an item
;;      naming an imported function — which lives in the EXPORTING module's
;;      class — resolved to nothing and became a null slot.
;;
;; All three fail the same way at the call site: "uninitialized element". The
;; message names the table slot, not the missing binding, so each of these
;; reads as a table bug rather than a naming one.

(module
  (func (export "ef0") (result i32) (i32.const 100))
  (func (export "ef1") (result i32) (i32.const 101))
  (func (export "ef2") (result i32) (i32.const 102))
)
(register "a")

(module
  (type $ret (func (result i32)))

  ;; Indices 0,1 — imported and UNNAMED, the case that had no binding.
  (import "a" "ef0" (func (result i32)))
  (import "a" "ef1" (func (result i32)))
  ;; Index 2 — imported and NAMED; the control that always worked.
  (import "a" "ef2" (func $named (result i32)))

  ;; Indices 3,4 — defined here.
  (func $own3 (result i32) (i32.const 3))
  (func $own4 (result i32) (i32.const 4))

  (table $t 12 funcref)

  ;; An ACTIVE segment written entirely in numbers, mixing imported (0,1,2)
  ;; and defined (3,4) functions. Under the bug every slot from an import
  ;; stayed null while the defined ones landed — so a test that used only
  ;; defined functions would have passed.
  (elem (table $t) (i32.const 0) func 0 1 2 3 4)

  ;; A PASSIVE segment, also numeric, copied into the table later. Its items
  ;; span both classes, which is what one-class-per-segment could not express.
  (elem $p funcref (ref.func 4) (ref.func 0) (ref.func 3) (ref.func 2))

  (func (export "call") (param $i i32) (result i32)
    (call_indirect $t (type $ret) (local.get $i)))

  ;; The direct-call side of the same index space.
  (func (export "direct0") (result i32) (call 0))
  (func (export "direct2") (result i32) (call 2))
  (func (export "direct_named") (result i32) (call $named))
  (func (export "direct4") (result i32) (call 4))

  (func (export "init_passive")
    (table.init $t $p (i32.const 5) (i32.const 0) (i32.const 4)))
  (func (export "drop_passive") (elem.drop $p))
  ;; A ZERO-length init: legal on a dropped segment, and the case an early
  ;; "segment dropped" return got wrong.
  (func (export "init_passive_zero")
    (table.init $t $p (i32.const 5) (i32.const 0) (i32.const 0)))
)

;; ── the active segment ─────────────────────────────────────────────────
(assert_return (invoke "call" (i32.const 0)) (i32.const 100))
(assert_return (invoke "call" (i32.const 1)) (i32.const 101))
(assert_return (invoke "call" (i32.const 2)) (i32.const 102))
(assert_return (invoke "call" (i32.const 3)) (i32.const 3))
(assert_return (invoke "call" (i32.const 4)) (i32.const 4))

;; ── direct calls by index ──────────────────────────────────────────────
(assert_return (invoke "direct0") (i32.const 100))
(assert_return (invoke "direct2") (i32.const 102))
(assert_return (invoke "direct_named") (i32.const 102))
(assert_return (invoke "direct4") (i32.const 4))

;; ── the passive segment ────────────────────────────────────────────────
;; Slot 5 was never written by the active segment, so it starts null.
(assert_trap (invoke "call" (i32.const 5)) "uninitialized element")
(invoke "init_passive")
(assert_return (invoke "call" (i32.const 5)) (i32.const 4))
(assert_return (invoke "call" (i32.const 6)) (i32.const 100))
(assert_return (invoke "call" (i32.const 7)) (i32.const 3))
(assert_return (invoke "call" (i32.const 8)) (i32.const 102))

;; ── a dropped segment is an EMPTY one, not an error ────────────────────
;; The spec drops the payload and leaves the segment in place: a ZERO-length
;; copy off a dropped segment SUCCEEDS and only a non-zero one traps. Both
;; halves matter — returning early on "dropped" made the zero-length case trap,
;; and having no check at all would let the non-zero case read freed data.
(invoke "drop_passive")
(assert_return (invoke "init_passive_zero"))
(assert_trap (invoke "init_passive") "out of bounds table access")
;; …and dropping twice is not itself an error.
(invoke "drop_passive")
