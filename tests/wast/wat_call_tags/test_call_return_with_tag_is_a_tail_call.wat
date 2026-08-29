;; vybe-test: wast/wat_call_tags/test_call_return_with_tag_is_a_tail_call
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; Design §Instructions: "`call_return_with_tag $call_tag : [ti* funcref] -> [to*]`
;; … TAIL calls the given `funcref` with the specified call tag".
;;
;; ⛔ `test_call_return_with_tag` proves only that the RESULT is right — one
;; call, one frame. An ordinary call passes it identically. The tail-ness is the
;; whole reason the instruction exists (the proposal cites the Tail Call
;; proposal for it) and nothing exercised it.
;;
;; 200_000 self-calls through `call_return_with_tag`. A genuine tail call reuses
;; the frame and answers 7; an ordinary call accumulates 200_000 frames and
;; exhausts the stack. The depth is the assertion.
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t (canon) (param i32) (result i32))

  (func $loop (param i32) (result i32)
    local.get 0
    i32.eqz
    if
      i32.const 7
      return
    end
    local.get 0
    i32.const 1
    i32.sub
    ref.func $loop
    call_return_with_tag $t
  )

  (func (export "_start")
    i32.const 200000
    call $loop
    call $log
  )
)
