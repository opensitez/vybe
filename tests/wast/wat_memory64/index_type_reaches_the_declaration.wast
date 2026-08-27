;; vybe-test: wast/wat_memory64/index_type_reaches_the_declaration
;; vybe-test-mode: run
;;
;; memory64 adds NO opcodes. It adds an INDEX TYPE to the declaration —
;; `(memory i64 …)`, `(table i64 …)` — and every address, count and size
;; operand of the ops that touch them is read at that width instead of i32.
;;
;; The VM has always had the machinery: `chunk.memory_is_64` / `table_is_64`,
;; and `pop_mem_index` / `pop_table_count` widening off them. Only the BINARY
;; reader ever filled them in. `StmtKind::MemoryDecl` / `TableDecl` carried no
;; index type at all, so a module written in TEXT declared a 64-bit memory and
;; got a 32-bit one — silently, because an i32 answer compares equal to the
;; i64 the test expects at every value small enough to be written by hand.
;; That is why the four cases that were already here passed while none of the
;; spec's own memory64 files did.
;;
;; The grammar rejected `(table i64 …)` outright, so the table half never even
;; parsed. Both halves are covered below, at values where the width is
;; OBSERVABLE: a size larger than the default memory, and a table index that
;; must survive as an i64 all the way into `call_indirect`.

(module
  ;; `$a` is a 64-bit memory of 3 pages; `memory.size` on it answers i64.
  (memory i64 3)

  (func (export "size") (result i64) (memory.size))
  (func (export "grow") (param $n i64) (result i64) (memory.grow (local.get $n)))

  ;; i64 ADDRESSES — the operand must not be truncated or read as an i32.
  (func (export "store_at") (param $at i64) (param $v i32)
    (i32.store (local.get $at) (local.get $v)))
  (func (export "load_at") (param $at i64) (result i32)
    (i32.load (local.get $at)))

  ;; An address beyond page 0 proves the memory really is 3 pages and that a
  ;; 64-bit address reaches it.
  (func (export "store_page2") (i32.store (i64.const 0x20000) (i32.const 0xabc)))
  (func (export "load_page2") (result i32) (i32.load (i64.const 0x20000)))

  ;; memory.fill / memory.copy take i64 counts on a 64-bit memory.
  (func (export "fill")
    (memory.fill (i64.const 64) (i32.const 9) (i64.const 4)))
  (func (export "copy")
    (memory.copy (i64.const 128) (i64.const 64) (i64.const 4)))
  (func (export "byte_at") (param $at i64) (result i32)
    (i32.load8_u (local.get $at)))
)

(assert_return (invoke "size") (i64.const 3))

(invoke "store_at" (i64.const 0) (i32.const 7))
(assert_return (invoke "load_at" (i64.const 0)) (i32.const 7))

(invoke "store_page2")
(assert_return (invoke "load_page2") (i32.const 0xabc))
;; …and page 0 is untouched by a write two pages up.
(assert_return (invoke "load_at" (i64.const 0)) (i32.const 7))

(invoke "fill")
(assert_return (invoke "byte_at" (i64.const 64)) (i32.const 9))
(assert_return (invoke "byte_at" (i64.const 67)) (i32.const 9))
(assert_return (invoke "byte_at" (i64.const 68)) (i32.const 0))

(invoke "copy")
(assert_return (invoke "byte_at" (i64.const 128)) (i32.const 9))
(assert_return (invoke "byte_at" (i64.const 131)) (i32.const 9))

;; `memory.grow` answers the OLD size, as an i64.
(assert_return (invoke "grow" (i64.const 2)) (i64.const 3))
(assert_return (invoke "size") (i64.const 5))

;; An address past the end still traps — widening the operand must not widen
;; the bound.
(assert_trap (invoke "load_at" (i64.const 0x50000)) "out of bounds memory access")

;; ── The `(memory i64 (data …))` abbreviation ─────────────────────────
;; The index type sits before the inline data, outside `mem_type`.
(module
  (memory i64 (data "\11\22\33"))
  (func (export "b") (param $at i64) (result i32) (i32.load8_u (local.get $at))))
(assert_return (invoke "b" (i64.const 0)) (i32.const 0x11))
(assert_return (invoke "b" (i64.const 2)) (i32.const 0x33))

;; ── Tables ──────────────────────────────────────────────────────────
;; `(table i64 …)` did not parse at all. Its size, its grow result and the
;; index `call_indirect` reads are all i64.
;;
;; `call_indirect` read that index with `as_f64()`, and `i64.const` lowers to
;; `Literal::BigInt` — whose `as_f64` is NaN. The trap it produced said
;; "table index NaN out of bounds", which is the shape to look for if this
;; regresses.
(module
  (type $ret_i32 (func (result i32)))

  (table $t i64 4 funcref)
  (elem (table $t) (i64.const 1) func $one $two)

  (func $one (result i32) (i32.const 111))
  (func $two (result i32) (i32.const 222))

  (func (export "size") (result i64) (table.size $t))
  (func (export "grow") (param $n i64) (result i64)
    (table.grow $t (ref.null func) (local.get $n)))
  (func (export "call") (param $i i64) (result i32)
    (call_indirect $t (type $ret_i32) (local.get $i)))
  (func (export "get_is_null") (param $i i64) (result i32)
    (ref.is_null (table.get $t (local.get $i))))
  (func (export "set_one") (param $i i64)
    (table.set $t (local.get $i) (ref.func $one)))
  (func (export "fill_null") (param $at i64) (param $n i64)
    (table.fill $t (local.get $at) (ref.null func) (local.get $n)))
)

(assert_return (invoke "size") (i64.const 4))

;; The i64 index must survive into the call — this is the case that trapped.
(assert_return (invoke "call" (i64.const 1)) (i32.const 111))
(assert_return (invoke "call" (i64.const 2)) (i32.const 222))

(assert_return (invoke "get_is_null" (i64.const 0)) (i32.const 1))
(assert_return (invoke "get_is_null" (i64.const 1)) (i32.const 0))

(invoke "set_one" (i64.const 3))
(assert_return (invoke "call" (i64.const 3)) (i32.const 111))

(invoke "fill_null" (i64.const 3) (i64.const 1))
(assert_return (invoke "get_is_null" (i64.const 3)) (i32.const 1))

(assert_return (invoke "grow" (i64.const 2)) (i64.const 4))
(assert_return (invoke "size") (i64.const 6))

;; Out of range traps rather than wrapping.
(assert_trap (invoke "call" (i64.const 6)) "undefined element")

;; ── A 32-bit declaration is still 32-bit ────────────────────────────
;; The control: if the index type were being applied unconditionally, or read
;; off the wrong declaration, these answer i64 instead.
(module
  (memory 1)
  (table 2 funcref)
  (func (export "msize") (result i32) (memory.size))
  (func (export "tsize") (result i32) (table.size)))
(assert_return (invoke "msize") (i32.const 1))
(assert_return (invoke "tsize") (i32.const 2))
