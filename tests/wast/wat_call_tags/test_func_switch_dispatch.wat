;; vybe-test: wast/wat_call_tags/test_func_switch_dispatch
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "`func_switch` is a new way of defining of functions […] specifying
;; essentially a switch statement that calls a `$func` if the given call tag
;; matches the corresponding `$call_tag`."
;;
;; ONE funcref answering TWO tags differently — the mechanism behind the
;; proposal's interface-method dispatch: one funcref per descriptor slot,
;; switching on the tag.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $get (param i32) (result i32))
  (call_tag $set (param i32) (result i32))

  (func $do_get (param i32) (result i32)
    local.get 0
    i32.const 1
    i32.add
  )
  (func $do_set (param i32) (result i32)
    local.get 0
    i32.const 100
    i32.add
  )

  (func_switch $slot
    (case $get $do_get)
    (case $set $do_set)
  )

  (func (export "_start")
    i32.const 5
    ref.func $slot
    call_with_tag $get
    call $log

    i32.const 5
    ref.func $slot
    call_with_tag $set
    call $log
  )
)
