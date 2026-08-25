;; vybe-test: wast/wat_call_tags/test_call_tag_same_signature_distinct
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; The property a structural type system cannot express. `$a` and `$b` have the
;; SAME wasm signature `(param i32) (result i32)` — after GC type
;; canonicalisation that is one type — yet each answers only its own tag, so the
;; two calling conventions stay apart. This is why call tags exist rather than
;; richer function types.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $ta (param i32) (result i32))
  (call_tag $tb (param i32) (result i32))

  (func $a (param i32) (result i32) (call_tag $ta)
    local.get 0
    i32.const 10
    i32.add
  )
  (func $b (param i32) (result i32) (call_tag $tb)
    local.get 0
    i32.const 20
    i32.add
  )

  (func (export "_start")
    i32.const 1
    ref.func $a
    call_with_tag $ta
    call $log

    i32.const 1
    ref.func $b
    call_with_tag $tb
    call $log
  )
)
