;; vybe-test: wast/wast_script_assert_malformed/assert_malformed_lexical_and_immediate_shapes
;; vybe-test-mode: run
;;
;; WAT lexes a maximal run of idchars as ONE token. Our grammar instead matches
;; an instruction name against a list and treats `_` as a digit, so the
;; SEPARATION the spec requires is simply not represented: `i32.const0` parsed
;; as `i32.const` then `0`, `i32.load32` as `i32.load` then `32`, and `_100` as
;; the number 100.
;;
;; Each of these was an `assert_malformed` in the spec's own files that passed
;; vacuously — the module parsed and nothing said so. They are checked over the
;; quoted TEXT, which runs for `assert_malformed` alone and cannot affect
;; ordinary compilation.
;;
;; The controls at the end matter as much as the cases: a check that rejects
;; too much makes a malformed assertion pass for the WRONG reason, which hides
;; real leniency instead of reporting it.

;; ── Token separation ────────────────────────────────────────────────
(assert_malformed (module quote "(func (drop (i32.const0)))") "unknown operator")
(assert_malformed (module quote "(func br 0drop)") "unknown operator")
(assert_malformed (module quote "(memory 1)(func (param i32) (result i32) (i32.load32 (local.get 0)))") "unknown operator")
(assert_malformed (module quote "(memory 1)(func (param i32) (result i32) (i32.load32_u (local.get 0)))") "unknown operator")
(assert_malformed (module quote "(memory 1)(func (param i32) (i32.store32 (local.get 0) (i32.const 0)))") "unknown operator")
(assert_malformed (module quote "(memory 1)(func (param i32) (f32.store64 (local.get 0) (f64.const 0)))") "unknown operator")

;; ── Number literals ─────────────────────────────────────────────────
(assert_malformed (module quote "(global i32 (i32.const _100))") "unknown operator")
(assert_malformed (module quote "(global i32 (i32.const +_100))") "unknown operator")
(assert_malformed (module quote "(global i32 (i32.const 99_))") "unknown operator")
(assert_malformed (module quote "(global i32 (i32.const 1__000))") "unknown operator")
(assert_malformed (module quote "(global i32 (i32.const _0x100))") "unknown operator")
(assert_malformed (module quote "(func (i32.const 0x) drop)") "unknown operator")
(assert_malformed (module quote "(func (i32.const 1x) drop)") "unknown operator")
(assert_malformed (module quote "(func (i32.const 0xg) drop)") "unknown operator")

;; A `const` MUST carry its literal, and it must be representable.
(assert_malformed (module quote "(func (i32.const) drop)") "unexpected token")
(assert_malformed (module quote "(func (f64.const) drop)") "unexpected token")
(assert_malformed (module quote "(func (result f32) (f32.const 0x1p128))") "constant out of range")
(assert_malformed (module quote "(func (result f64) (f64.const 0x1p1024))") "constant out of range")

;; `nan:canonical` / `nan:arithmetic` are RESULT patterns, not values.
(assert_malformed (module quote "(func (result f32) (f32.const nan:arithmetic))") "unexpected token")
(assert_malformed (module quote "(func (result f64) (f64.const nan:canonical))") "unexpected token")

;; ── Identifiers ─────────────────────────────────────────────────────
(assert_malformed (module quote "(func $)") "empty identifier")
(assert_malformed (module quote "(func $\"\")") "empty identifier")
(assert_malformed (module quote "(func $ \"a\")") "empty identifier")
(assert_malformed (module quote "(func $\"a\nb\")") "empty identifier")
(assert_malformed (module quote "(func $\"\\ef\")") "malformed UTF-8")

;; ── Alignment: the memarg carries the EXPONENT ──────────────────────
(assert_malformed (module quote "(memory 0) (func (drop (i32.load align=0 (i32.const 0))))") "alignment")
(assert_malformed (module quote "(memory 0) (func (drop (i64.load align=7 (i32.const 0))))") "alignment")

;; ── Block types: `(type)? (param)* (result)*`, and no NAMED params ──
(assert_malformed (module quote "(func (i32.const 0) (block (result i32) (param i32)))") "unexpected token")
(assert_malformed (module quote "(func (i32.const 0) (block (param $x i32) (drop)))") "unexpected token")
(assert_malformed
  (module quote "(type $sig (func))" "(func (block (type $sig) (result i32) (i32.const 0)) (unreachable))")
  "inline function type")

;; ── `catch` / `catch_all` are try_table CLAUSES, not instructions ───
(assert_malformed (module quote "(func (catch_all))") "unexpected token")
(assert_malformed (module quote "(tag $e) (func (catch $e))") "unexpected token")

;; ── SIMD immediates ─────────────────────────────────────────────────
;; A lane INDEX is one unsigned byte; a lane VALUE must fit its shape.
(assert_malformed (module quote "(func (result i32) (i8x16.extract_lane_s -1 (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)))") "unexpected token")
(assert_malformed (module quote "(func (result i32) (i8x16.extract_lane_s 256 (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)))") "i8 constant out of range")
(assert_malformed (module quote "(func (result i32) (i8x16.extract_lane_s (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0)))") "unexpected token")
(assert_malformed (module quote "(func (v128.const i8x16 256 256 256 256 256 256 256 256 256 256 256 256 256 256 256 256) drop)") "constant out of range")
(assert_malformed (module quote "(func (v128.const i8x16 -129 -129 -129 -129 -129 -129 -129 -129 -129 -129 -129 -129 -129 -129 -129 -129) drop)") "constant out of range")
(assert_malformed (module quote "(func (v128.const) drop)") "unexpected token")
(assert_malformed (module quote "(func (v128.const i32x4 0 0 0) drop)") "wrong number of lane literals")
(assert_malformed (module quote "(func (param v128) (result v128) (i8x16.shuffle 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 (local.get 0) (local.get 0)))") "invalid lane length")

;; ── Controls: none of the above may reject well-formed text ─────────
;; If any of these were taken as malformed, the assertion would PASS for the
;; wrong reason and the check would be hiding leniency rather than finding it.
(module (memory 1) (func (drop (i32.load8_u align=1 (i32.const 0)))))
(module (memory 1) (func (drop (i32.load offset=7 align=4 (i32.const 0)))))
(module (memory 1) (func (drop (i64.load32_u (i32.const 0)))))
(module (func (result i32) (i32.const 1_000_000)))
(module (func (result i32) (i32.const 0x1_0000)))
(module (func (result f64) (f64.const 0x1.921fb54442d18p+1)))
(module (func (result f64) (f64.const 1e10)))
(module (func (result f32) (f32.const nan:0x200000)))
(module (func (result f32) (f32.const inf)))
(module (func (result f32) (f32.const 0x1p127)))
(module (func $a (result i32) (i32.const 1)))
(module (func (i32.const 0) (block (param i32) (result i32)) (drop)))
(module (type $s (func (param i32) (result i32)))
        (func (i32.const 0) (block (type $s) (param i32) (result i32)) (drop)))
(module (func (result v128) (v128.const i32x4 1 2 3 4)))
(module (func (result i32) (i8x16.extract_lane_s 15 (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))))
(module (func (param v128) (result v128) (i8x16.shuffle 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 (local.get 0) (local.get 0))))
(module (tag $e) (func (try_table (catch_all 0))))
