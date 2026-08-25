;; vybe-test: wast/wat_component/test_a_map_key_type_is_read
;; hand-written against proposals/component-model/design/mvp/CanonicalABI.md:2181
;;   case MapType(k, v) : return ListType(despecialize(TupleType([k, v])))
;;
;; ▶▶ THE KEY TYPE IS READ, AND ITS WIDTH DECIDES THE ELEMENT LAYOUT.
;;
;; This file used to pin a REFUSAL: it put `char` in the key position and read
;; the message back, because `char` was unrepresentable and naming it proved
;; the key had been lowered. `char` is implemented now, and every spelling a
;; `keytype` allows is supported, so no refusal is available — the claim has to
;; be made positively instead.
;;
;; `map<u8, u8>` despecialises to `list<record{"0": u8, "1": u8}>`. Both fields
;; align to 1, so an entry is TWO bytes: key at +0, value at +1.
;;
;;   100: 0x07   the key
;;   101: 0x2A   the value, 42
;;
;; The callee answers `value * 100 + key` = 4207, so BOTH halves must survive
;; and in the right order.
;;
;;   4207  correct
;;    742  key and value read in the wrong order
;;
;; ⛔ WHAT THIS FILE DOES **NOT** PROVE, MEASURED RATHER THAN ASSUMED: that the
;; key's WIDTH drives the element layout. Changing the type to `map<u8, u8>` →
;; `map<u32, u8>` still passes, and that is not a defect — it is a limit of any
;; probe that reads raw bytes at low offsets.
;;
;; Under `map<u8,u8>` an entry is 2 bytes, key at +0 and value at +1. Under
;; `map<u32,u8>` it is 8, key at +0 and value at +4. But the key is LITTLE
;; ENDIAN, so loading four bytes and storing them back into the copy reproduces
;; the same four bytes — byte +1 of the copy is source byte 101 under BOTH
;; layouts. The two are byte-identical exactly where this probe looks, and
;; moving the probe to a high offset only reads past the copy.
;;
;; So the assert below pins PRESENCE and ORDER of the key, which is what the
;; original bug destroyed. Pinning the key's WIDTH needs a probe that reads
;; where the two layouts genuinely differ, and this is not it. Said plainly
;; here rather than left implied by a green result.
;;
;; ⛔ THE ORIGINAL BUG THIS FILE EXISTS FOR: a map's key is `Rule::keytype`,
;; its OWN atomic rule, because the spec restricts a key to the primitives with
;; a total ordering (no float, no `error-context`). The walk collected children
;; through a helper filtering for `Rule::valtype`, so it skipped the key, read
;; the VALUE type as the key, then refused with `map: no value type` — naming
;; the half of the source that was PRESENT.
;;
;; That bug could not show while `map` refused outright, and would not have
;; shown afterwards either: `list<record{k,v}>` and a one-field record both
;; flatten to the same `(ptr, len)` pair, so only the ELEMENT LAYOUT differs —
;; which is exactly what this file now measures.

(component
  (core module $m
    (memory (export "mem") 1)
    (data (i32.const 100) "\07\2A")
    (global $bump (mut i32) (i32.const 1024))
    (func (export "realloc") (param i32 i32 i32 i32) (result i32)
      (local $p i32)
      (local.set $p
        (i32.and
          (i32.add (global.get $bump) (i32.sub (local.get 2) (i32.const 1)))
          (i32.sub (i32.const 0) (local.get 2))))
      (global.set $bump (i32.add (local.get $p) (local.get 3)))
      (local.get $p))
    (func (export "probe") (param i32 i32) (result i32)
      (i32.add
        (i32.mul (i32.load8_u (i32.add (local.get 0) (i32.const 1)))
                 (i32.const 100))
        (i32.load8_u (local.get 0)))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "probe" (core func $c))
  (alias core export $mi "realloc" (core func $r))

  (type $ft (func (param "m" (map u8 u8)) (result u32)))
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

(assert_return (invoke "get") (i32.const 4207))
