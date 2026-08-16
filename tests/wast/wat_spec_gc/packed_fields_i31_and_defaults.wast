;; vybe-test: wast/wat_spec_gc/packed_fields_i31_and_defaults
;; vybe-test-mode: run
;;
;; From the aggregate and i31 instructions in core/exec/instructions.rst and
;; the storage-type rules in core/syntax/types.rst.
;;
;; The load-bearing rules:
;;
;; * A PACKED field (i8/i16) stores only its own width, and the width is
;;   restored on read by the ACCESSOR, not the field: `struct.get_s` sign-
;;   extends and `struct.get_u` zero-extends the identical stored bits. A
;;   packed field written 0xFF reads back as -1 or 255 depending purely on
;;   which accessor asked. An implementation that stores the full i32 and
;;   ignores the packing answers 255 for both.
;; * `*.new_default` initialises numeric fields to 0 and reference fields to
;;   null — it does not leave them undefined.
;; * `ref.i31` keeps the low 31 bits, so it is lossy for any i32 with bit 31
;;   set, and `i31.get_s` sign-extends from bit 30 rather than bit 31. That
;;   makes `i31.get_s(ref.i31(-1))` equal -1 while `i31.get_u` of the same
;;   reference is 2147483647.
;; * `ref.eq` is reference identity, so two separately allocated structs with
;;   equal contents are NOT equal.

(module
  (type $packed (struct (field (mut i8)) (field (mut i16)) (field (mut i32))))
  (type $refs (struct (field (mut anyref)) (field (mut i32))))
  (type $bytes (array (mut i8)))
  (type $words (array (mut i32)))

  (func (export "p_get_s8") (param i32) (result i32)
    (struct.get_s $packed 0 (struct.new $packed (local.get 0) (i32.const 0) (i32.const 0))))
  (func (export "p_get_u8") (param i32) (result i32)
    (struct.get_u $packed 0 (struct.new $packed (local.get 0) (i32.const 0) (i32.const 0))))
  (func (export "p_get_s16") (param i32) (result i32)
    (struct.get_s $packed 1 (struct.new $packed (i32.const 0) (local.get 0) (i32.const 0))))
  (func (export "p_get_u16") (param i32) (result i32)
    (struct.get_u $packed 1 (struct.new $packed (i32.const 0) (local.get 0) (i32.const 0))))
  ;; An unpacked i32 field keeps every bit and has only a plain accessor.
  (func (export "p_get32") (param i32) (result i32)
    (struct.get $packed 2 (struct.new $packed (i32.const 0) (i32.const 0) (local.get 0))))

  ;; Packing applies on WRITE as well as on construction.
  (func (export "set_then_get_s") (param i32) (result i32)
    (local $s (ref $packed))
    (local.set $s (struct.new_default $packed))
    (struct.set $packed 0 (local.get $s) (local.get 0))
    (struct.get_s $packed 0 (local.get $s)))

  ;; new_default: numeric fields 0, reference fields null.
  (func (export "default_i32") (result i32)
    (struct.get $refs 1 (struct.new_default $refs)))
  (func (export "default_ref_is_null") (result i32)
    (ref.is_null (struct.get $refs 0 (struct.new_default $refs))))

  ;; ── i31 ───────────────────────────────────────────────────────────────
  (func (export "i31_s") (param i32) (result i32) (i31.get_s (ref.i31 (local.get 0))))
  (func (export "i31_u") (param i32) (result i32) (i31.get_u (ref.i31 (local.get 0))))

  ;; ── arrays ────────────────────────────────────────────────────────────
  (func (export "arr_default") (result i32)
    (array.get_u $bytes (array.new_default $bytes (i32.const 4)) (i32.const 0)))
  (func (export "arr_len") (result i32)
    (array.len (array.new_default $bytes (i32.const 7))))
  (func (export "arr_packed_s") (param i32) (result i32)
    (local $a (ref $bytes))
    (local.set $a (array.new_default $bytes (i32.const 2)))
    (array.set $bytes (local.get $a) (i32.const 0) (local.get 0))
    (array.get_s $bytes (local.get $a) (i32.const 0)))
  (func (export "arr_packed_u") (param i32) (result i32)
    (local $a (ref $bytes))
    (local.set $a (array.new_default $bytes (i32.const 2)))
    (array.set $bytes (local.get $a) (i32.const 0) (local.get 0))
    (array.get_u $bytes (local.get $a) (i32.const 0)))
  ;; array.new fills every element with the same value.
  (func (export "arr_filled") (param i32) (result i32)
    (array.get $words (array.new $words (local.get 0) (i32.const 3)) (i32.const 2)))
  (func (export "arr_oob") (param i32) (result i32)
    (array.get $words (array.new $words (i32.const 5) (i32.const 3)) (local.get 0)))
  (func (export "arr_fill_then_read") (result i32)
    (local $a (ref $words))
    (local.set $a (array.new_default $words (i32.const 5)))
    (array.fill $words (local.get $a) (i32.const 1) (i32.const 9) (i32.const 3))
    (i32.add (array.get $words (local.get $a) (i32.const 0))
             (array.get $words (local.get $a) (i32.const 3))))
  (func (export "arr_fill_edge") (result i32)
    (local $a (ref $words))
    (local.set $a (array.new_default $words (i32.const 5)))
    (array.fill $words (local.get $a) (i32.const 1) (i32.const 9) (i32.const 3))
    (array.get $words (local.get $a) (i32.const 4)))

  ;; ── null aggregate access traps ───────────────────────────────────────
  (func (export "null_struct_get") (result i32)
    (struct.get $packed 2 (ref.null $packed)))
  (func (export "null_array_get") (result i32)
    (array.get $words (ref.null $words) (i32.const 0)))
  (func (export "null_array_len") (result i32)
    (array.len (ref.null $words)))

  ;; ── ref.eq is identity, not structural equality ───────────────────────
  (func (export "same_struct_eq") (result i32)
    (local $s (ref $refs))
    (local.set $s (struct.new_default $refs))
    (ref.eq (local.get $s) (local.get $s)))
  (func (export "equal_structs_eq") (result i32)
    (ref.eq (struct.new_default $refs) (struct.new_default $refs)))
  ;; i31 references with the same value ARE the same reference — an i31 is not
  ;; allocated, its value IS its identity.
  (func (export "same_i31_eq") (result i32)
    (ref.eq (ref.i31 (i32.const 7)) (ref.i31 (i32.const 7))))
  (func (export "diff_i31_eq") (result i32)
    (ref.eq (ref.i31 (i32.const 7)) (ref.i31 (i32.const 8))))
)

;; ── Packed i8: the accessor decides the sign ────────────────────────────
(assert_return (invoke "p_get_s8" (i32.const 255)) (i32.const -1))
(assert_return (invoke "p_get_u8" (i32.const 255)) (i32.const 255))
(assert_return (invoke "p_get_s8" (i32.const 127)) (i32.const 127))
(assert_return (invoke "p_get_u8" (i32.const 127)) (i32.const 127))
(assert_return (invoke "p_get_s8" (i32.const 128)) (i32.const -128))
(assert_return (invoke "p_get_u8" (i32.const 128)) (i32.const 128))
;; Bits above the field's width are discarded on the way IN.
(assert_return (invoke "p_get_u8" (i32.const 0x1234ff)) (i32.const 255))
(assert_return (invoke "p_get_s8" (i32.const 0x1234ff)) (i32.const -1))

;; ── Packed i16 ─────────────────────────────────────────────────────────
(assert_return (invoke "p_get_s16" (i32.const 65535)) (i32.const -1))
(assert_return (invoke "p_get_u16" (i32.const 65535)) (i32.const 65535))
(assert_return (invoke "p_get_s16" (i32.const 32768)) (i32.const -32768))
(assert_return (invoke "p_get_u16" (i32.const 32768)) (i32.const 32768))
(assert_return (invoke "p_get_u16" (i32.const 0x12340000)) (i32.const 0))

;; An unpacked field is not truncated at all.
(assert_return (invoke "p_get32" (i32.const -1)) (i32.const -1))
(assert_return (invoke "p_get32" (i32.const 0x12345678)) (i32.const 0x12345678))

;; Packing happens on struct.set too, not only on construction.
(assert_return (invoke "set_then_get_s" (i32.const 255)) (i32.const -1))
(assert_return (invoke "set_then_get_s" (i32.const 1)) (i32.const 1))

;; ── new_default ────────────────────────────────────────────────────────
(assert_return (invoke "default_i32") (i32.const 0))
(assert_return (invoke "default_ref_is_null") (i32.const 1))
(assert_return (invoke "arr_default") (i32.const 0))
(assert_return (invoke "arr_len") (i32.const 7))

;; ── i31: 31 bits, sign taken from bit 30 ──────────────────────────────
(assert_return (invoke "i31_s" (i32.const 7)) (i32.const 7))
(assert_return (invoke "i31_u" (i32.const 7)) (i32.const 7))
(assert_return (invoke "i31_s" (i32.const -1)) (i32.const -1))
(assert_return (invoke "i31_u" (i32.const -1)) (i32.const 2147483647))
(assert_return (invoke "i31_s" (i32.const 0x40000000)) (i32.const -1073741824))
(assert_return (invoke "i31_u" (i32.const 0x40000000)) (i32.const 1073741824))
;; Bit 31 is discarded, so these two agree despite differing as i32.
(assert_return (invoke "i31_u" (i32.const 0x3fffffff)) (i32.const 1073741823))
(assert_return (invoke "i31_s" (i32.const 0x3fffffff)) (i32.const 1073741823))
(assert_return (invoke "i31_u" (i32.const -2147483648)) (i32.const 0))

;; ── arrays ─────────────────────────────────────────────────────────────
(assert_return (invoke "arr_packed_s" (i32.const 200)) (i32.const -56))
(assert_return (invoke "arr_packed_u" (i32.const 200)) (i32.const 200))
(assert_return (invoke "arr_filled" (i32.const 42)) (i32.const 42))
;; array.fill covers exactly [start, start+n).
(assert_return (invoke "arr_fill_then_read") (i32.const 9))
(assert_return (invoke "arr_fill_edge") (i32.const 0))
;; Bounds: last valid index is len-1, and the index is unsigned.
(assert_return (invoke "arr_oob" (i32.const 2)) (i32.const 5))
(assert_trap (invoke "arr_oob" (i32.const 3)) "out of bounds array access")
(assert_trap (invoke "arr_oob" (i32.const -1)) "out of bounds array access")

;; ── null aggregate access ──────────────────────────────────────────────
(assert_trap (invoke "null_struct_get") "null structure reference")
(assert_trap (invoke "null_array_get") "null array reference")
(assert_trap (invoke "null_array_len") "null array reference")

;; ── ref.eq ─────────────────────────────────────────────────────────────
(assert_return (invoke "same_struct_eq") (i32.const 1))
(assert_return (invoke "equal_structs_eq") (i32.const 0))
(assert_return (invoke "same_i31_eq") (i32.const 1))
(assert_return (invoke "diff_i31_eq") (i32.const 0))
