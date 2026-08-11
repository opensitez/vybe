;; vybe-test: wast/wat_lexical/float_invalid_formats
;; origin: languages/wast/tests/wast/test_wat_lexical.rs
;; vybe-test-mode: compile-fail

;; A float literal must have a `num` BEFORE the dot — `.5` has none.
;;
;; This test previously also asserted that `1.` and `0x1.5` were malformed.
;; Both are VALID per the spec and per the official suite, which declares them
;; as plain (non-assert_malformed) modules:
;;   const.wast:47          (module (func (f32.const 0123456789.) drop))
;;   float_literals.wast:105 (f32.const 0xa0_ff.f141_a59a)
;; The spec's `float ::= num '.' frac?` allows an EMPTY fraction, and its
;; `hexfloat` requires the `p` exponent only when there is no dot.
(module (global f32 (f32.const .5)))
