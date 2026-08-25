;; vybe-test: wast/wat_call_tags/test_call_indirect_with_tag
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; "`call_indirect_with_tag $table $call_tag : [ti* i32] -> [to*]` is shorthand
;; for `(call_with_tag $call_tag (table.get $table))`."
;;
;; Same resolution as the direct form, with the funcref fetched from a table.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (table 1 funcref)
  (elem (i32.const 0) $triple)

  (call_tag $t (canon) (param i32) (result i32))

  (func $triple (param i32) (result i32)
    local.get 0
    i32.const 3
    i32.mul
  )

  (func (export "_start")
    i32.const 5
    i32.const 0
    call_indirect_with_tag 0 $t
    call $log
  )
)
