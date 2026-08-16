;; vybe-test: wast/wat_spec_reftypes/reftype_abbreviations_are_abstract_types
;; vybe-test-mode: run
;;
;; From core/syntax/types.rst §2.3.4, where each of these is defined as pure
;; shorthand for a nullable reference to an ABSTRACT heap type:
;;
;;   anyref ≡ (ref null any)        funcref   ≡ (ref null func)
;;   eqref ≡ (ref null eq)          externref ≡ (ref null extern)
;;   i31ref ≡ (ref null i31)        nullref   ≡ (ref null none)
;;   structref ≡ (ref null struct)  nullfuncref   ≡ (ref null nofunc)
;;   arrayref ≡ (ref null array)    nullexternref ≡ (ref null noextern)
;;
;; Nothing in the corpus tested `ref.test` against any of these spellings — all
;; four files that mention `ref.test` at all use a concrete `$t` or a bare heap
;; type. The abbreviations therefore fell through the emitter's name resolution
;; to its CONCRETE branch, which reserved a module type slot named "funcref"
;; and emitted the immediate as that slot's INDEX. The instruction then tested
;; the value against an unrelated declared type and could never be true:
;; `(ref.test funcref (ref.func $f))` disassembled to `ref.test 2` and answered
;; 0. Every assertion below is 1 where the spec says the type matches.
;;
;; Each hierarchy is tested separately because they do not mix: `func`,
;; `extern` and `any` are disjoint, and testing across them is a validation
;; error rather than a false result.

(module
  (type $thunk (func (result i32)))
  (type $point (struct (field i32)))
  (type $ints (array (mut i32)))
  (elem declare func $ten)
  (func $ten (type $thunk) (i32.const 10))

  ;; ── func hierarchy ────────────────────────────────────────────────────
  (func (export "func_is_funcref") (result i32)
    (ref.test funcref (ref.func $ten)))
  ;; The long form must agree with the abbreviation exactly.
  (func (export "func_is_ref_null_func") (result i32)
    (ref.test (ref null func) (ref.func $ten)))
  (func (export "func_is_nonnull_func") (result i32)
    (ref.test (ref func) (ref.func $ten)))
  ;; A real function is not the BOTTOM of its hierarchy.
  (func (export "func_is_nullfuncref") (result i32)
    (ref.test nullfuncref (ref.func $ten)))
  ;; A null funcref matches the nullable spellings but not the non-null one.
  (func (export "null_is_funcref") (result i32)
    (ref.test funcref (ref.null func)))
  (func (export "null_is_nonnull_func") (result i32)
    (ref.test (ref func) (ref.null func)))
  (func (export "null_is_nullfuncref") (result i32)
    (ref.test nullfuncref (ref.null func)))

  ;; ── extern hierarchy ──────────────────────────────────────────────────
  (func (export "null_is_externref") (result i32)
    (ref.test externref (ref.null extern)))
  (func (export "null_is_nullexternref") (result i32)
    (ref.test nullexternref (ref.null extern)))
  (func (export "extern_of_i31_is_externref") (result i32)
    (ref.test externref (extern.convert_any (ref.i31 (i32.const 7)))))
  ;; A converted i31 is a real extern, so it is NOT the bottom type.
  (func (export "extern_of_i31_is_nullexternref") (result i32)
    (ref.test nullexternref (extern.convert_any (ref.i31 (i32.const 7)))))

  ;; ── any hierarchy: i31 ────────────────────────────────────────────────
  (func (export "i31_is_anyref") (result i32)
    (ref.test anyref (ref.i31 (i32.const 7))))
  (func (export "i31_is_eqref") (result i32)
    (ref.test eqref (ref.i31 (i32.const 7))))
  (func (export "i31_is_i31ref") (result i32)
    (ref.test i31ref (ref.i31 (i32.const 7))))
  ;; i31 is a sibling of struct and array, not a subtype.
  (func (export "i31_is_structref") (result i32)
    (ref.test structref (ref.i31 (i32.const 7))))
  (func (export "i31_is_arrayref") (result i32)
    (ref.test arrayref (ref.i31 (i32.const 7))))
  (func (export "i31_is_nullref") (result i32)
    (ref.test nullref (ref.i31 (i32.const 7))))

  ;; ── any hierarchy: struct ─────────────────────────────────────────────
  (func (export "struct_is_anyref") (result i32)
    (ref.test anyref (struct.new $point (i32.const 1))))
  (func (export "struct_is_eqref") (result i32)
    (ref.test eqref (struct.new $point (i32.const 1))))
  (func (export "struct_is_structref") (result i32)
    (ref.test structref (struct.new $point (i32.const 1))))
  (func (export "struct_is_arrayref") (result i32)
    (ref.test arrayref (struct.new $point (i32.const 1))))
  (func (export "struct_is_i31ref") (result i32)
    (ref.test i31ref (struct.new $point (i32.const 1))))

  ;; ── any hierarchy: array ──────────────────────────────────────────────
  (func (export "array_is_anyref") (result i32)
    (ref.test anyref (array.new $ints (i32.const 0) (i32.const 2))))
  (func (export "array_is_arrayref") (result i32)
    (ref.test arrayref (array.new $ints (i32.const 0) (i32.const 2))))
  (func (export "array_is_structref") (result i32)
    (ref.test structref (array.new $ints (i32.const 0) (i32.const 2))))
  (func (export "array_is_eqref") (result i32)
    (ref.test eqref (array.new $ints (i32.const 0) (i32.const 2))))

  ;; ── the null of the `any` hierarchy ───────────────────────────────────
  (func (export "nullany_is_anyref") (result i32)
    (ref.test anyref (ref.null any)))
  (func (export "nullany_is_nullref") (result i32)
    (ref.test nullref (ref.null any)))
  (func (export "nullany_is_nonnull_any") (result i32)
    (ref.test (ref any) (ref.null any)))
  (func (export "nullany_is_i31ref") (result i32)
    (ref.test i31ref (ref.null any)))
)

;; ── func ─────────────────────────────────────────────────────────────────
(assert_return (invoke "func_is_funcref") (i32.const 1))
(assert_return (invoke "func_is_ref_null_func") (i32.const 1))
(assert_return (invoke "func_is_nonnull_func") (i32.const 1))
(assert_return (invoke "func_is_nullfuncref") (i32.const 0))
(assert_return (invoke "null_is_funcref") (i32.const 1))
(assert_return (invoke "null_is_nonnull_func") (i32.const 0))
(assert_return (invoke "null_is_nullfuncref") (i32.const 1))

;; ── extern ───────────────────────────────────────────────────────────────
(assert_return (invoke "null_is_externref") (i32.const 1))
(assert_return (invoke "null_is_nullexternref") (i32.const 1))
(assert_return (invoke "extern_of_i31_is_externref") (i32.const 1))
(assert_return (invoke "extern_of_i31_is_nullexternref") (i32.const 0))

;; ── i31 ──────────────────────────────────────────────────────────────────
(assert_return (invoke "i31_is_anyref") (i32.const 1))
(assert_return (invoke "i31_is_eqref") (i32.const 1))
(assert_return (invoke "i31_is_i31ref") (i32.const 1))
(assert_return (invoke "i31_is_structref") (i32.const 0))
(assert_return (invoke "i31_is_arrayref") (i32.const 0))
(assert_return (invoke "i31_is_nullref") (i32.const 0))

;; ── struct ───────────────────────────────────────────────────────────────
(assert_return (invoke "struct_is_anyref") (i32.const 1))
(assert_return (invoke "struct_is_eqref") (i32.const 1))
(assert_return (invoke "struct_is_structref") (i32.const 1))
(assert_return (invoke "struct_is_arrayref") (i32.const 0))
(assert_return (invoke "struct_is_i31ref") (i32.const 0))

;; ── array ────────────────────────────────────────────────────────────────
(assert_return (invoke "array_is_anyref") (i32.const 1))
(assert_return (invoke "array_is_arrayref") (i32.const 1))
(assert_return (invoke "array_is_structref") (i32.const 0))
(assert_return (invoke "array_is_eqref") (i32.const 1))

;; ── null of `any` ────────────────────────────────────────────────────────
(assert_return (invoke "nullany_is_anyref") (i32.const 1))
(assert_return (invoke "nullany_is_nullref") (i32.const 1))
(assert_return (invoke "nullany_is_nonnull_any") (i32.const 0))
(assert_return (invoke "nullany_is_i31ref") (i32.const 1))
