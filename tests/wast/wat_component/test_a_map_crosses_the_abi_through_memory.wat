;; vybe-test: wast/wat_component/test_a_map_crosses_the_abi_through_memory
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md
;;   §Despecialization (:2181):
;;     case MapType(k, v) : return ListType(despecialize(TupleType([k, v])))
;;   and §Storing / §Loading for `list` and `string`.
;;
;; ▶▶ THE FIRST NON-SCALAR COMPONENT VALUE TO CROSS THE ABI. Every other
;; component test in this directory passes integers, which the flat
;; representation carries in core parameters and which therefore never touch
;; linear memory, `realloc`, or a layout rule. A `map<string, u32>` touches all
;; three at once:
;;
;;   map<string,u32> ↦ list<tuple<string,u32>> ↦ list<record{"0":string,"1":u32}>
;;
;; so the value is a (ptr, len) pair addressing records of
;; [str_ptr:u32][str_len:u32][value:u32] — 12 bytes, align 4 — each of whose
;; strings is itself a (ptr, len) into memory.
;;
;; The source map is one entry, `"hi" -> 7`, laid down by a data segment at 100:
;;
;;   100: str_ptr = 112     104: str_len = 2     108: value = 7
;;   112: "hi"
;;
;; `canon lower` LIFTS that out of memory, the lifted function LOWERS it back
;; through `realloc`, and the core callee reads the copy. So both directions of
;; the memory ABI run on one value.
;;
;; ⛔ THE ANSWER IS COMPOSITE ON PURPOSE: 7 * 1000 + 'h'. Reading only the
;; value would leave the KEY unproven, and the key is the half that was
;; silently dropped — `(map …)`'s key is `Rule::keytype`, its own atomic rule,
;; and the walk collected `Rule::valtype` children only. See
;; `test_a_map_key_type_is_read`. So:
;;
;;   7104  correct                     7105  read 'i', off by one in the string
;;   2104  read str_len as the value   1104  read str_ptr as the value
;;
;; ⛔ THE DECLARED `realloc` IS NOT WHAT ALLOCATES. The module exports one and
;; the canon rows name it, but `CanonOpts::realloc` is written by the walker,
;; carried through the compiler and **read by nothing** — canonical lowering
;; allocates from the VM's own marshalling bump global
;; (`dispatch.rs canon_bump_start`). Measured: give the module's allocator a
;; floor of 50000 and the list still arrives near 65536. `CanonOpts::memory`
;; is write-only in the same way. Both are recorded in cmplan.md §Known
;; deviations; this file asserts only what is true today, which is that a COPY
;; happened somewhere above the data segment.

(component
  (core module $m
    (memory (export "mem") 1)
    (data (i32.const 100) "\70\00\00\00\02\00\00\00\07\00\00\00hi")
    (global $bump (mut i32) (i32.const 1024))
    (func (export "realloc") (param i32 i32 i32 i32) (result i32)
      (local $p i32)
      ;; align the bump pointer up to `align`, then hand out `new_size`.
      (local.set $p
        (i32.and
          (i32.add (global.get $bump) (i32.sub (local.get 2) (i32.const 1)))
          (i32.sub (i32.const 0) (local.get 2))))
      (global.set $bump (i32.add (local.get $p) (local.get 3)))
      (local.get $p))
    (func (export "probe") (param i32 i32) (result i32)
      ;; ⛔ THE LIST MUST BE A COPY, NOT THE SOURCE. The caller hands `canon
      ;; lower` the literal address 100, which is where the data segment
      ;; already sits — so a pointer passed through unchanged would read the
      ;; ORIGINAL bytes and answer 7104 too, and the test would pass without
      ;; the ABI having moved anything. Every allocator in play hands out well
      ;; above 1024 and the source sits below it, so the ADDRESS separates
      ;; them: a source-address read shows up as 107104.
      (i32.add
        ;; 100000 iff the list was NOT copied above the bump floor
        (i32.mul (i32.lt_u (local.get 0) (i32.const 1024)) (i32.const 100000))
        (i32.add
          ;; the u32 value of entry 0 — field "1" of the record, at offset 8
          (i32.mul (i32.load (i32.add (local.get 0) (i32.const 8)))
                   (i32.const 1000))
          ;; the first byte of entry 0's KEY — field "0" is a (ptr, len) at 0
          (i32.load8_u (i32.load (local.get 0)))))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "probe" (core func $c))
  (alias core export $mi "realloc" (core func $r))

  (type $ft (func (param "m" (map string u32)) (result u32)))
  (canon lift  (core func $c) (memory (core memory 0)) (realloc (core func $r))
               (func $f (type $ft)))
  (canon lower (func $f)      (memory (core memory 0)) (realloc (core func $r))
               (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32 i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 100) (i32.const 1))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 7104))
