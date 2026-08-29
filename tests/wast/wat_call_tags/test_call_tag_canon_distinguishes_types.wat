;; vybe-test: wast/wat_call_tags/test_call_tag_canon_distinguishes_types
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;; vybe-test-mode: run-fail
;;
;; Design §Tags: "`call_tag.canon $functype` derives the canonical call tag OF
;; TYPE `$functype`". Two different functypes are two different tags — the whole
;; point of the canonical form is that it is per-TYPE.
;;
;; ⛔ IT USED TO BE PER-ARITY. `canonical_call_tags` was keyed
;; `HashMap<(u8,u8), u32>` and `CallTagDef` carried only `params`/`results`, so
;; `[i32]->[i32]` and `[f64]->[f64]` interned to ONE tag and this module was
;; ACCEPTED: an i32-shaped funcref answered the f64 canonical tag.
;;
;; That mattered beyond this test. `call_indirect $table $functype` is shorthand
;; for `call_with_tag (call_tag.canon $functype)`, so with an arity key the
;; Security property — "a funcref called under a convention it does not handle
;; STOPS, rather than being called anyway under the wrong shape" — was only
;; ARITY-safety, and an i32-shaped and an f64-shaped callee stayed
;; interchangeable through the front door.
;;
;; CONTROLS, both of which must keep passing: `test_call_tag_canonical_default`
;; (a func DOES answer the canonical tag of its own type) and
;; `test_call_indirect_undeclared_func_still_works` (plain `call_indirect` still
;; reaches an undeclared func — its immediates carry counts, not types, so it
;; asks for the shape tag).
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  ;; Same arity, different types — therefore two distinct canonical tags.
  (call_tag $ci (canon) (param i32) (result i32))
  (call_tag $cf (canon) (param f64) (result f64))

  (func $takes_i32 (param i32) (result i32)
    local.get 0
  )

  (func (export "_start")
    i32.const 5
    ref.func $takes_i32
    call_with_tag $cf
    call $log
  )
)
