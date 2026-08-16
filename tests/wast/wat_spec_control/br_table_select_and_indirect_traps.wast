;; vybe-test: wast/wat_spec_control/br_table_select_and_indirect_traps
;; vybe-test-mode: run
;;
;; From core/exec/instructions.rst, "Control Instructions".
;;
;; The rules being pinned:
;;
;; * `br_table l* lN`: the index selects from `l*`, and ANY index at or past
;;   the table's length uses the default `lN`. The index is read UNSIGNED, so
;;   -1 is 4294967295 and takes the default rather than the last entry.
;; * `br l` to a BLOCK label jumps to its end; `br l` to a LOOP label jumps to
;;   its start. The label's identity decides the direction, not the branch.
;; * `select` picks the FIRST operand when the condition is non-zero. It is a
;;   plain instruction, not a conditional: both value operands are already on
;;   the stack, so neither is skipped.
;; * `call_indirect x y` has three distinct trap conditions, and they are
;;   genuinely different failures: index past the table, a null entry, and a
;;   type that does not match the declared one.
;; * `unreachable` always traps.

(module
  (type $ret_i32 (func (result i32)))
  (type $takes_i32 (func (param i32) (result i32)))

  (table 4 funcref)
  (func $ten (type $ret_i32) (i32.const 10))
  (func $twenty (type $ret_i32) (i32.const 20))
  (func $inc (type $takes_i32) (i32.add (local.get 0) (i32.const 1)))
  ;; Slots 0,1 hold `$ret_i32` functions, slot 2 a `$takes_i32` one,
  ;; slot 3 is left NULL on purpose.
  (elem (i32.const 0) $ten $twenty $inc)

  (func (export "br_table") (param i32) (result i32)
    (block $default
      (block $c2
        (block $c1
          (block $c0
            (br_table $c0 $c1 $c2 $default (local.get 0)))
          (return (i32.const 100)))
        (return (i32.const 101)))
      (return (i32.const 102)))
    (i32.const 999))

  ;; Every entry of the table names the SAME label: the index then cannot
  ;; change the outcome, which is the spec's degenerate case.
  (func (export "br_table_same") (param i32) (result i32)
    (block $out
      (br_table $out $out $out (local.get 0))
      (return (i32.const 1)))
    (i32.const 2))

  ;; `br 0` inside a loop restarts it; inside a block it exits. Same
  ;; instruction, opposite direction, decided by the label.
  (func (export "loop_counts") (param i32) (result i32)
    (local $acc i32)
    (block $exit
      (loop $again
        (br_if $exit (i32.eqz (local.get 0)))
        (local.set $acc (i32.add (local.get $acc) (local.get 0)))
        (local.set 0 (i32.sub (local.get 0) (i32.const 1)))
        (br $again)))
    (local.get $acc))

  (func (export "select") (param i32 i32 i32) (result i32)
    (select (local.get 0) (local.get 1) (local.get 2)))
  (func (export "select_f64") (param f64 f64 i32) (result f64)
    (select (local.get 0) (local.get 1) (local.get 2)))

  (func (export "call_indirect") (param i32) (result i32)
    (call_indirect (type $ret_i32) (local.get 0)))
  (func (export "call_indirect_arg") (param i32 i32) (result i32)
    (call_indirect (type $takes_i32) (local.get 1) (local.get 0)))
  (func (export "unreachable") (result i32)
    (unreachable))
  (func (export "unreachable_after_value") (result i32)
    (i32.const 1)
    (unreachable))
)

;; ── br_table: in-range indices, then everything else takes the default ────
(assert_return (invoke "br_table" (i32.const 0)) (i32.const 100))
(assert_return (invoke "br_table" (i32.const 1)) (i32.const 101))
(assert_return (invoke "br_table" (i32.const 2)) (i32.const 102))
;; Exactly at the table length — the first index that is out of range.
(assert_return (invoke "br_table" (i32.const 3)) (i32.const 999))
(assert_return (invoke "br_table" (i32.const 4)) (i32.const 999))
(assert_return (invoke "br_table" (i32.const 1000)) (i32.const 999))
;; Read UNSIGNED: -1 is 4294967295, so it is out of range, not the last entry.
(assert_return (invoke "br_table" (i32.const -1)) (i32.const 999))
(assert_return (invoke "br_table" (i32.const -2147483648)) (i32.const 999))

;; When every entry names one label the index cannot matter.
(assert_return (invoke "br_table_same" (i32.const 0)) (i32.const 2))
(assert_return (invoke "br_table_same" (i32.const 1)) (i32.const 2))
(assert_return (invoke "br_table_same" (i32.const 99)) (i32.const 2))
(assert_return (invoke "br_table_same" (i32.const -1)) (i32.const 2))

;; ── br to a loop label restarts; br_if to a block label exits ────────────
(assert_return (invoke "loop_counts" (i32.const 0)) (i32.const 0))
(assert_return (invoke "loop_counts" (i32.const 1)) (i32.const 1))
(assert_return (invoke "loop_counts" (i32.const 4)) (i32.const 10))
(assert_return (invoke "loop_counts" (i32.const 100)) (i32.const 5050))

;; ── select takes the FIRST operand when the condition is non-zero ────────
(assert_return (invoke "select" (i32.const 1) (i32.const 2) (i32.const 1)) (i32.const 1))
(assert_return (invoke "select" (i32.const 1) (i32.const 2) (i32.const 0)) (i32.const 2))
;; Any non-zero condition counts, including negative and the sign bit alone.
(assert_return (invoke "select" (i32.const 1) (i32.const 2) (i32.const -1)) (i32.const 1))
(assert_return (invoke "select" (i32.const 1) (i32.const 2) (i32.const -2147483648)) (i32.const 1))
;; Values pass through bit-exact, including a signed zero and a NaN.
(assert_return (invoke "select_f64" (f64.const -0) (f64.const 1) (i32.const 1)) (f64.const -0))
(assert_return (invoke "select_f64" (f64.const nan) (f64.const 1) (i32.const 1))
               (f64.const nan:canonical))
(assert_return (invoke "select_f64" (f64.const nan) (f64.const 1) (i32.const 0)) (f64.const 1))
(assert_return (invoke "select_f64" (f64.const inf) (f64.const -inf) (i32.const 0))
               (f64.const -inf))

;; ── call_indirect: the ordinary case, then its three distinct traps ──────
(assert_return (invoke "call_indirect" (i32.const 0)) (i32.const 10))
(assert_return (invoke "call_indirect" (i32.const 1)) (i32.const 20))
(assert_return (invoke "call_indirect_arg" (i32.const 2) (i32.const 41)) (i32.const 42))

;; Index past the end of the table.
(assert_trap (invoke "call_indirect" (i32.const 4)) "undefined element")
(assert_trap (invoke "call_indirect" (i32.const 100)) "undefined element")
;; The index is unsigned here too.
(assert_trap (invoke "call_indirect" (i32.const -1)) "undefined element")
;; In range, but the slot was never initialised.
(assert_trap (invoke "call_indirect" (i32.const 3)) "uninitialized element")
;; In range and non-null, but the wrong type — slot 2 takes an i32 parameter.
(assert_trap (invoke "call_indirect" (i32.const 2)) "indirect call type mismatch")
;; ...and the mirror image: a `$takes_i32` call landing on a no-parameter entry.
(assert_trap (invoke "call_indirect_arg" (i32.const 0) (i32.const 1))
             "indirect call type mismatch")

;; ── unreachable ─────────────────────────────────────────────────────────
(assert_trap (invoke "unreachable") "unreachable")
;; Values already on the stack do not save it.
(assert_trap (invoke "unreachable_after_value") "unreachable")
