;; vybe-test: wast/wat_call_tags/test_call_return_with_tag_tail_through_fallback
;; origin: proposals/call-tags/proposals/call-tags/Overview.md
;;
;; Two clauses meeting, and the interesting case is where they meet:
;;   §Instructions — `call_return_with_tag` "TAIL calls the given funcref"
;;   §Functions    — a funcref that does not handle the tag reaches "the
;;                   fall-back handler of the call tag", which is "(tail) called"
;;
;; So a MISS under `call_return_with_tag` must stay in tail position: the frame
;; belongs to neither the refused funcref nor the handler, and must already be
;; gone before either is chosen. `test_call_return_with_tag_is_a_tail_call`
;; covers only the HANDLED path; nothing covered the miss.
;;
;; `$only` declares $t1, so every $t2 call misses it and lands in `$h`, which
;; recurses 100_000 times through the same miss. Tail position is preserved iff
;; the frame is dropped BEFORE dispatch decides handled-vs-fallback — which is
;; why the pop lives at the instruction and not inside either branch.
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))

  (call_tag $t1 (param i32) (result i32))
  (call_tag $t2 (param i32) (result i32) (fallback $h))

  ;; Handles $t1 only — so it refuses every $t2 call below.
  (func $only (param i32) (result i32) (call_tag $t1)
    local.get 0
  )

  (func $h (param i32) (param externref) (result i32)
    local.get 0
    i32.eqz
    if
      i32.const 7
      return
    end
    local.get 0
    i32.const 1
    i32.sub
    ref.func $only
    call_return_with_tag $t2
  )

  ;; The tail call needs a containing function whose results MATCH the tag's —
  ;; `_start` returns nothing, and the subtype rule this suite also tests
  ;; rejects a tail call out of it. (It caught this test's first draft.)
  (func $run (param i32) (result i32)
    local.get 0
    ref.func $only
    call_return_with_tag $t2
  )

  (func (export "_start")
    i32.const 100000
    call $run
    call $log
  )
)
