;; vybe-test: wast/wat_memory_addressing/the_memarg_offset_is_not_a_signed_add
;; vybe-test-mode: run
;;
;; WASM computes a load/store's effective address in UNBOUNDED arithmetic over
;; the address read as UNSIGNED (§4.4.7):
;;
;;     ea = unsigned(i) + memarg.offset      -- and ea + N/8 > |mem| traps
;;
;; The walker used to FOLD `offset=N` into the address expression, producing an
;; AST `addr + N`. That is a SIGNED add on a signed value, and the two disagree
;; on exactly the inputs the spec singles out: at address -1 the spec computes
;; 4294967295 + 25 and traps, while the fold computed 24, stayed in bounds and
;; happily returned a byte. Every `-1` assertion in `address.wast` is this.
;;
;; The offset now rides in the instruction's MEMARG — the VM's `effective_addr`
;; already did the unsigned widen and a saturating add — carried across as an
;; `@@off<N>` name suffix, the same emitter channel `@@mem<N>` uses.
;;
;; ⛔ The fold was ALSO blind in the PLAIN spelling. There the address is not an
;; argument at all — it is on the enclosing block's stack — so there was no slot
;; to fold into and the offset was DROPPED. A folded-only test cannot see that:
;; it is the same asymmetry that made a folded `br_on_cast` repro understate its
;; bug. Every case below is written BOTH ways.

(module
  (memory 1)
  (data (i32.const 0) "abcdefghijklmnopqrstuvwxyz")

  ;; ── folded spelling ──────────────────────────────────────────────────
  (func (export "f_load8") (param $i i32) (result i32)
    (i32.load8_u offset=2 align=1 (local.get $i)))
  (func (export "f_load16") (param $i i32) (result i32)
    (i32.load16_u offset=2 align=2 (local.get $i)))
  (func (export "f_load32") (param $i i32) (result i32)
    (i32.load offset=2 align=4 (local.get $i)))
  (func (export "f_load64") (param $i i32) (result i64)
    (i64.load offset=2 (local.get $i)))
  (func (export "f_store8") (param $i i32) (param $v i32)
    (i32.store8 offset=2 (local.get $i) (local.get $v)))

  ;; ── plain spelling — the address is on the stack, never an argument ──
  (func (export "p_load8") (param $i i32) (result i32)
    local.get $i
    i32.load8_u offset=2 align=1)
  (func (export "p_load16") (param $i i32) (result i32)
    local.get $i
    i32.load16_u offset=2 align=2)
  (func (export "p_load32") (param $i i32) (result i32)
    local.get $i
    i32.load offset=2 align=4)
  (func (export "p_load64") (param $i i32) (result i64)
    local.get $i
    i64.load offset=2)
  (func (export "p_store8") (param $i i32) (param $v i32)
    local.get $i
    local.get $v
    i32.store8 offset=2)

  ;; A memarg offset written in HEX. The whole of `align.wast` writes its
  ;; offsets this way, and a plain `parse::<u64>()` read `0x008` as 0 —
  ;; silently, and only ever in the safe direction (a load that should have
  ;; been displaced read the byte at the base instead).
  (func (export "hex_offset") (param $i i32) (result i32)
    (i32.load8_u offset=0x19 (local.get $i)))
  (func (export "hex_offset_plain") (param $i i32) (result i32)
    local.get $i
    i32.load8_u offset=0x19)

  ;; `align=` is a pure hint: it is validated, but the semantics do not depend
  ;; on it, so an under-aligned load reads exactly the same bytes.
  (func (export "align_is_a_hint") (param $i i32) (result i32)
    (i32.load offset=1 align=1 (local.get $i)))

  ;; The largest offset a 32-bit memarg can carry. Legal to write on any
  ;; address; it is the SUM that decides, and 0 + 4294967295 is out of bounds.
  (func (export "huge_offset") (param $i i32) (result i32)
    (i32.load8_u offset=4294967295 (local.get $i)))
)

;; ── the offset actually displaces (folded and plain agree) ─────────────
(assert_return (invoke "f_load8" (i32.const 0)) (i32.const 99))   ;; 'c'
(assert_return (invoke "p_load8" (i32.const 0)) (i32.const 99))
(assert_return (invoke "f_load8" (i32.const 3)) (i32.const 102))  ;; 'f'
(assert_return (invoke "p_load8" (i32.const 3)) (i32.const 102))
(assert_return (invoke "f_load16" (i32.const 0)) (i32.const 25699))    ;; 'cd'
(assert_return (invoke "p_load16" (i32.const 0)) (i32.const 25699))
(assert_return (invoke "f_load32" (i32.const 0)) (i32.const 1717920867)) ;; 'cdef'
(assert_return (invoke "p_load32" (i32.const 0)) (i32.const 1717920867))
(assert_return (invoke "f_load64" (i32.const 0)) (i64.const 7667774633883821155))
(assert_return (invoke "p_load64" (i32.const 0)) (i64.const 7667774633883821155))

(assert_return (invoke "hex_offset" (i32.const 0)) (i32.const 122))       ;; 0x19 = 25, 'z'
(assert_return (invoke "hex_offset_plain" (i32.const 0)) (i32.const 122))
(assert_return (invoke "align_is_a_hint" (i32.const 0)) (i32.const 1701077858)) ;; 'bcde'

;; A store displaces by the same offset, and the load that reads it back must
;; agree — a fix that moved only the load half would pass every assertion above.
(invoke "f_store8" (i32.const 30) (i32.const 65))
(assert_return (invoke "f_load8" (i32.const 30)) (i32.const 65))
(invoke "p_store8" (i32.const 40) (i32.const 66))
(assert_return (invoke "p_load8" (i32.const 40)) (i32.const 66))
;; …and it really landed at base+2, not at the base.
(assert_return (invoke "f_load8" (i32.const 28)) (i32.const 0))
(assert_return (invoke "f_load8" (i32.const 38)) (i32.const 0))

;; ── THE CASE THE SIGNED FOLD GOT WRONG ─────────────────────────────────
;; -1 is 4294967295 unsigned, so every one of these is far out of bounds. The
;; fold computed 1 and returned the byte at address 1.
(assert_trap (invoke "f_load8" (i32.const -1)) "out of bounds memory access")
(assert_trap (invoke "p_load8" (i32.const -1)) "out of bounds memory access")
(assert_trap (invoke "f_load16" (i32.const -1)) "out of bounds memory access")
(assert_trap (invoke "p_load16" (i32.const -1)) "out of bounds memory access")
(assert_trap (invoke "f_load32" (i32.const -1)) "out of bounds memory access")
(assert_trap (invoke "p_load32" (i32.const -1)) "out of bounds memory access")
(assert_trap (invoke "f_load64" (i32.const -1)) "out of bounds memory access")
(assert_trap (invoke "p_load64" (i32.const -1)) "out of bounds memory access")
;; A STORE at -1 must trap too, and must not have written anything at 1.
(assert_trap (invoke "f_store8" (i32.const -1) (i32.const 88)) "out of bounds memory access")
(assert_trap (invoke "p_store8" (i32.const -1) (i32.const 88)) "out of bounds memory access")

;; ⛔ Do not "improve" the two lines above into an `assert_return` that reads
;; back address 1: the store trapped, so nothing was written anywhere, and an
;; assertion that the byte at 1 is unchanged passes just as well against the
;; OLD folded behaviour if the fold happened to write the same value.

;; ── the offset is added, never wrapped ────────────────────────────────
;; 0 + 4294967295 is out of bounds on a one-page memory, and so is 1 + it. A
;; 32-bit wrapping add would have made address 1 come back to 0 and succeed.
(assert_trap (invoke "huge_offset" (i32.const 0)) "out of bounds memory access")
(assert_trap (invoke "huge_offset" (i32.const 1)) "out of bounds memory access")

;; ── the last byte in range, and the first byte past it ─────────────────
;; One page is 65536 bytes; `offset=2` on a load8 makes 65533 the last legal
;; address and 65534 the first trapping one.
(assert_return (invoke "f_load8" (i32.const 65533)) (i32.const 0))
(assert_return (invoke "p_load8" (i32.const 65533)) (i32.const 0))
(assert_trap (invoke "f_load8" (i32.const 65534)) "out of bounds memory access")
(assert_trap (invoke "p_load8" (i32.const 65534)) "out of bounds memory access")

;; ── an `offset=` field is a uN — `offset=-1` is not a memarg at all ────
;; The lexer cannot build a memarg from a signed number, so the reference
;; implementation reports the whole instruction as an unknown operator rather
;; than anything about offsets.
(assert_malformed
  (module quote "(memory 1)(func (drop (i32.load offset=-1 (i32.const 0))))")
  "unknown operator"
)
(assert_malformed
  (module quote "(memory 1)(func (i32.store offset=-1 (i32.const 0) (i32.const 0)))")
  "unknown operator"
)
;; ⛔ …and the CONTROL, because a malformed check that rejects too much makes
;; the assertions above pass for the wrong reason. An offset TOO LARGE for a
;; 32-bit memory lexes perfectly well — the spec calls that INVALID, not
;; malformed — and the ordinary spellings must still compile and run.
(module
  (memory 1)
  (func (export "still_compiles") (result i32)
    (i32.load8_u offset=0 (i32.const 0)))
)
(assert_return (invoke "still_compiles") (i32.const 0))
