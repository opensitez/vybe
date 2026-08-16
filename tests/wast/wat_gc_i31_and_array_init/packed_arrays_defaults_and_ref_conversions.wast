;; vybe-test: wast/wat_gc_i31_and_array_init/packed_arrays_defaults_and_ref_conversions
;; vybe-test-mode: run
;;
;; The GC corners that occurred at most once in the run corpus: `array.get_s`,
;; `array.fill`, `array.init_data`, `array.new_elem`, `struct.new_default`,
;; `ref.i31` with a negative payload, `any.convert_extern` /
;; `extern.convert_any`, and `br_on_non_null`.
;;
;; What each one needs to be tested AT:
;;
;;   * `array.get_s` vs `array.get_u` on a PACKED array (`i8`/`i16`) holding a
;;     byte with the high bit set — the only place the two differ. On an
;;     unpacked `i32` array there is no `_s`/`_u` at all.
;;   * `struct.new_default` / `array.new_default`: the defaults must be the
;;     ZEROES of each field type, not whatever the allocator left.
;;   * `array.fill` / `array.init_data`: a sub-range, so the untouched elements
;;     on both sides are part of the assertion — a fill that ran long or short
;;     is otherwise invisible.
;;   * `ref.i31` / `i31.get_s` / `i31.get_u`: a payload with bit 30 set, where
;;     the signed and unsigned reads differ by 2^31.
;;   * `any.convert_extern` / `extern.convert_any`: a ROUND TRIP that ends in a
;;     value read back, so an identity-losing conversion shows.
;;   * `br_on_non_null`: both directions, and the fact that it leaves the
;;     (non-null) reference on the stack for the target.
;;
;; Spec-format so `wasmtime wast` arbitrates every expectation.

(module
  (type $bytes (array (mut i8)))
  (type $words (array (mut i16)))
  (type $ints (array (mut i32)))
  (type $funcs (array (mut funcref)))
  (type $pt (struct (field $x (mut i32)) (field $y (mut f64))))
  (type $thunk (func (result i32)))

  (table $t 1 funcref)
  (data $d "\01\ff\7f\80")
  (elem $e func $ten $twenty)

  (func $ten (result i32) (i32.const 10))
  (func $twenty (result i32) (i32.const 20))

  ;; ── packed arrays: _s and _u differ, and only here ───────────────────
  (func (export "get_s_i8") (result i32)
    (local $a (ref $bytes))
    (local.set $a (array.new_default $bytes (i32.const 4)))
    (array.set $bytes (local.get $a) (i32.const 0) (i32.const 255))
    (array.get_s $bytes (local.get $a) (i32.const 0)))
  (func (export "get_u_i8") (result i32)
    (local $a (ref $bytes))
    (local.set $a (array.new_default $bytes (i32.const 4)))
    (array.set $bytes (local.get $a) (i32.const 0) (i32.const 255))
    (array.get_u $bytes (local.get $a) (i32.const 0)))
  (func (export "get_s_i16") (result i32)
    (local $a (ref $words))
    (local.set $a (array.new_default $words (i32.const 2)))
    (array.set $words (local.get $a) (i32.const 1) (i32.const 65535))
    (array.get_s $words (local.get $a) (i32.const 1)))
  (func (export "get_u_i16") (result i32)
    (local $a (ref $words))
    (local.set $a (array.new_default $words (i32.const 2)))
    (array.set $words (local.get $a) (i32.const 1) (i32.const 65535))
    (array.get_u $words (local.get $a) (i32.const 1)))
  ;; A packed store keeps only the low bits.
  (func (export "packed_store_truncates") (result i32)
    (local $a (ref $bytes))
    (local.set $a (array.new_default $bytes (i32.const 1)))
    (array.set $bytes (local.get $a) (i32.const 0) (i32.const 513))
    (array.get_u $bytes (local.get $a) (i32.const 0)))

  ;; ── defaults are zeroes ──────────────────────────────────────────────
  (func (export "array_new_default_is_zero") (result i32)
    (array.get $ints (array.new_default $ints (i32.const 4)) (i32.const 3)))
  (func (export "struct_new_default_i32") (result i32)
    (struct.get $pt $x (struct.new_default $pt)))
  (func (export "struct_new_default_f64") (result f64)
    (struct.get $pt $y (struct.new_default $pt)))

  ;; ── array.fill writes a SUB-RANGE ───────────────────────────────────
  (func (export "fill_middle") (result i32)
    (local $a (ref $ints))
    (local.set $a (array.new_default $ints (i32.const 5)))
    (array.fill $ints (local.get $a) (i32.const 1) (i32.const 7) (i32.const 3))
    ;; 0,7,7,7,0 packed into one number so every element is asserted at once
    (i32.add
      (i32.add
        (i32.mul (array.get $ints (local.get $a) (i32.const 0)) (i32.const 10000))
        (i32.mul (array.get $ints (local.get $a) (i32.const 1)) (i32.const 1000)))
      (i32.add
        (i32.add
          (i32.mul (array.get $ints (local.get $a) (i32.const 2)) (i32.const 100))
          (i32.mul (array.get $ints (local.get $a) (i32.const 3)) (i32.const 10)))
        (array.get $ints (local.get $a) (i32.const 4)))))

  ;; ── array.init_data copies from a data segment ──────────────────────
  ;; Segment is 01 FF 7F 80; copy the middle two into elements 1..2.
  (func (export "init_data_middle") (result i32)
    (local $a (ref $bytes))
    (local.set $a (array.new_default $bytes (i32.const 4)))
    (array.init_data $bytes $d
      (local.get $a) (i32.const 1) (i32.const 1) (i32.const 2))
    ;; element 0 untouched (0), 1 = 0xFF, 2 = 0x7F, 3 untouched (0)
    (i32.add
      (i32.add
        (i32.mul (array.get_u $bytes (local.get $a) (i32.const 0)) (i32.const 16777216))
        (i32.mul (array.get_u $bytes (local.get $a) (i32.const 1)) (i32.const 65536)))
      (i32.add
        (i32.mul (array.get_u $bytes (local.get $a) (i32.const 2)) (i32.const 256))
        (array.get_u $bytes (local.get $a) (i32.const 3)))))

  ;; ── array.new_elem builds an array of refs from an elem segment ──────
  ;; Reached through a table + `call_indirect` rather than `ref.cast` +
  ;; `call_ref`: a funcref produced by `array.new_elem` currently fails
  ;; `ref.test`/`ref.cast` against a concrete func type (it carries no type
  ;; identity, unlike one from `ref.func`). That is a separate defect with its
  ;; own repro; going through the table keeps THIS file about `array.new_elem`.
  (func (export "new_elem_length") (result i32)
    (array.len (array.new_elem $funcs $e (i32.const 0) (i32.const 2))))
  (func (export "new_elem_elements_are_not_null") (result i32)
    (local $a (ref $funcs))
    (local.set $a (array.new_elem $funcs $e (i32.const 0) (i32.const 2)))
    (i32.add
      (i32.mul (ref.is_null (array.get $funcs (local.get $a) (i32.const 0)))
               (i32.const 2))
      (ref.is_null (array.get $funcs (local.get $a) (i32.const 1)))))
  (func (export "new_elem_calls_second") (result i32)
    (local $a (ref $funcs))
    (local.set $a (array.new_elem $funcs $e (i32.const 0) (i32.const 2)))
    (table.set $t (i32.const 0) (array.get $funcs (local.get $a) (i32.const 1)))
    (call_indirect $t (type $thunk) (i32.const 0)))
  (func (export "new_elem_calls_first") (result i32)
    (local $a (ref $funcs))
    (local.set $a (array.new_elem $funcs $e (i32.const 0) (i32.const 2)))
    (table.set $t (i32.const 0) (array.get $funcs (local.get $a) (i32.const 0)))
    (call_indirect $t (type $thunk) (i32.const 0)))

  ;; ── i31: signed and unsigned reads of the same payload ───────────────
  ;; 0x40000000 has bit 30 set, so as a 31-bit signed value it is negative.
  (func (export "i31_get_s") (result i32)
    (i31.get_s (ref.i31 (i32.const 0x40000000))))
  (func (export "i31_get_u") (result i32)
    (i31.get_u (ref.i31 (i32.const 0x40000000))))
  ;; ref.i31 keeps the low 31 bits.
  (func (export "i31_truncates") (result i32)
    (i31.get_u (ref.i31 (i32.const 0x80000005))))

  ;; ── extern <-> any round trip must preserve the value ───────────────
  (func (export "extern_any_roundtrip") (result i32)
    (i31.get_s
      (ref.cast (ref i31)
        (any.convert_extern (extern.convert_any (ref.i31 (i32.const -7)))))))
  ;; A null survives the round trip as a null.
  (func (export "extern_any_roundtrip_null") (result i32)
    (ref.is_null (any.convert_extern (extern.convert_any (ref.null any)))))

  ;; ── br_on_non_null: branches WITH the reference, or falls through ────
  (func (export "br_on_non_null_taken") (result i32)
    (block $found (result (ref $thunk))
      (br_on_non_null $found (ref.func $twenty))
      (return (i32.const -1)))
    (call_ref $thunk))
  (func (export "br_on_non_null_not_taken") (result i32)
    (block $found (result (ref $thunk))
      (br_on_non_null $found (ref.null $thunk))
      (return (i32.const -1)))
    (call_ref $thunk))
)

;; ── packed arrays ────────────────────────────────────────────────────
(assert_return (invoke "get_s_i8") (i32.const -1))
(assert_return (invoke "get_u_i8") (i32.const 255))
(assert_return (invoke "get_s_i16") (i32.const -1))
(assert_return (invoke "get_u_i16") (i32.const 65535))
(assert_return (invoke "packed_store_truncates") (i32.const 1))

;; ── defaults ─────────────────────────────────────────────────────────
(assert_return (invoke "array_new_default_is_zero") (i32.const 0))
(assert_return (invoke "struct_new_default_i32") (i32.const 0))
(assert_return (invoke "struct_new_default_f64") (f64.const 0))

;; ── fill / init_data ─────────────────────────────────────────────────
(assert_return (invoke "fill_middle") (i32.const 7770))
(assert_return (invoke "init_data_middle") (i32.const 16744192))

;; ── new_elem ─────────────────────────────────────────────────────────
(assert_return (invoke "new_elem_length") (i32.const 2))
(assert_return (invoke "new_elem_elements_are_not_null") (i32.const 0))
(assert_return (invoke "new_elem_calls_first") (i32.const 10))
(assert_return (invoke "new_elem_calls_second") (i32.const 20))

;; ── i31 ──────────────────────────────────────────────────────────────
(assert_return (invoke "i31_get_s") (i32.const -1073741824))
(assert_return (invoke "i31_get_u") (i32.const 1073741824))
(assert_return (invoke "i31_truncates") (i32.const 5))

;; ── extern <-> any ───────────────────────────────────────────────────
(assert_return (invoke "extern_any_roundtrip") (i32.const -7))
(assert_return (invoke "extern_any_roundtrip_null") (i32.const 1))

;; ── br_on_non_null ───────────────────────────────────────────────────
(assert_return (invoke "br_on_non_null_taken") (i32.const 20))
(assert_return (invoke "br_on_non_null_not_taken") (i32.const -1))
