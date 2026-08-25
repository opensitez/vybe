;; vybe-test: wast/wat_call_tags/test_call_tag_export_import_identity
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "Similarly, one can import and export call tags."
;;
;; An imported tag must be the EXPORTER's entity, not a fresh local one — that
;; is the whole contract, since a funcref handling the export has to answer a
;; call made under the import.
;;
;; What this asserts: the provider exports `conv`; the consumer imports it as
;; `$mine` and dispatches under it to a func declaring the same imported tag.
;; If the import did not resolve, `call_with_tag $mine` would fail with
;; "undefined call tag" — the tag would not exist in the consumer at all.
;;
;; NOT asserted here, and deliberately: dispatch to the PROVIDER's function
;; across the boundary. That needs `ref.func` on an imported func, which
;; returns `undefined` in this front end regardless of call tags — verified
;; with a plain `ref.func` + `call_ref` on an import, no tags involved. That is
;; a wast linking gap, not a call-tag one, and faking around it here would make
;; this test pass for a reason it does not test.

(module $provider
  (call_tag $shared (export "conv") (param i32) (result i32))
)

(register "provider" $provider)

(module $consumer
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "provider" "conv" (call_tag $mine (param i32) (result i32)))

  (func $local (param i32) (result i32) (call_tag $mine)
    local.get 0
    i32.const 5
    i32.add
  )

  (func (export "_start")
    i32.const 1
    ref.func $local
    call_with_tag $mine
    call $log
  )
)
