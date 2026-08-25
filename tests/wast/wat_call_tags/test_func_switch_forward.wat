;; vybe-test: wast/wat_call_tags/test_func_switch_forward
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "If `Bar` extends another class `Baz`, each of `Bar`'s `funcref`s can forward
;; unhandled call tags to `Baz`'s corresponding `funcref`. This can be used to
;; reduce duplication […] and to support separate compilation."
;;
;; `$sub` handles only $own; a call under $inherited falls through its forward
;; to `$base`, which does handle it.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $own (param i32) (result i32))
  (call_tag $inherited (param i32) (result i32))

  (func $sub_impl (param i32) (result i32)
    local.get 0
    i32.const 1
    i32.add
  )
  (func $base_impl (param i32) (result i32)
    local.get 0
    i32.const 2
    i32.add
  )

  (func_switch $base
    (case $inherited $base_impl)
  )
  (func_switch $sub
    (case $own $sub_impl)
    (forward $base)
  )

  (func (export "_start")
    i32.const 10
    ref.func $sub
    call_with_tag $own
    call $log

    i32.const 10
    ref.func $sub
    call_with_tag $inherited
    call $log
  )
)
