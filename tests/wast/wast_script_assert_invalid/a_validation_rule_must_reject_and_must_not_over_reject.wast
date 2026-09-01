;; vybe-test: wast/wast_script_assert_invalid/a_validation_rule_must_reject_and_must_not_over_reject
;; vybe-test-mode: run
;;
;; ⛔ EVERY VALIDATION RULE HAS TWO HALVES, AND ONLY ONE OF THEM IS CHEAP.
;;
;; A rule that fires is easy to demonstrate: point it at the invalid module the
;; spec names. The half that actually costs something is the other one — the
;; VALID modules the rule must leave alone. Nothing in an `assert_invalid`
;; suite can show an over-fire, because every module there is invalid already;
;; a rule that rejected everything would score perfectly. So each rule below is
;; pinned from BOTH sides: the module the spec rejects, and the neighbouring
;; module it must still accept.
;;
;; ⛔ AND THE VALID HALF BELOW CANNOT, BY ITSELF, CATCH AN OVER-FIRE TODAY.
;; Validation runs only inside `assert_invalid` — a plain `(module …)` is
;; never checked — so a rule that rejects everything would still let these
;; modules build. Removing `(elem declare func $by_elem)` from the module
;; below leaves this file GREEN, and that is measured, not assumed. What the
;; valid half does pin is that these modules keep COMPILING AND RUNNING, and
;; it becomes a real over-fire guard the moment validation moves onto the
;; normal path. The over-fire check that bites right now is wrapping the
;; suite's valid modules in `assert_invalid` (`fpcheck.py`), which is how the
;; seven false positives named below were found.
;;
;; Three of these were written from the fixtures alone and would have been
;; wrong:
;;
;;   * `ref.func` — the declared set excludes function BODIES and the START
;;     function, but an INLINE EXPORT on a function declares it. Skipping the
;;     whole `func` field reads the rule off the two invalid fixtures and
;;     rejects `(func $f (export "a") (drop (ref.func $f)))`, which is valid.
;;
;;   * constant expressions — WASM 3.0 folds in extended-const, so
;;     `i32.add`/`sub`/`mul` ARE constant. The pre-3.0 set rejects
;;     `elem.wast:1062` and `data.wast:180`, both valid. `i32.ctz` is the line.
;;
;;   * `global.get` in a constant expression — how much of the global index
;;     space is visible depends on WHERE the expression sits, and the reason is
;;     instantiation order (§4.5.4): a table is allocated before any global is
;;     initialised, a global at index i sees `[0, i)`, and an elem/data offset
;;     runs after all of them. Reading that as one rule ("imports only")
;;     rejected SEVEN valid modules across global/elem/data.

;; ── a branch to the outermost label, and the rules around it ──────────────
(module
  ;; ref.func: declared by an elem segment, by a global, and by an inline
  ;; export — all three must be accepted.
  (func $by_elem)
  (func $by_global)
  (func $by_export (export "e"))
  (elem declare func $by_elem)
  (global funcref (ref.func $by_global))
  (func (export "refs") (result i32)
    (drop (ref.func $by_elem))
    (drop (ref.func $by_global))
    (drop (ref.func $by_export))
    (i32.const 1)
  )

  ;; A start function with the right type is fine, and reading a MUTABLE
  ;; global from a BODY is fine — it is only an initializer that may not.
  (global $mut (mut i32) (i32.const 0))
  (func $start_ok (global.set $mut (i32.const 3)))
  (start $start_ok)
  (func (export "mut") (result i32) (global.get $mut))

  ;; Lane immediates at the last legal index of each shape.
  (func (export "lanes") (result i32)
    (i32.add
      (i8x16.extract_lane_s 15 (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 7))
      (i32x4.extract_lane 3 (v128.const i32x4 0 0 0 9)))
  )
  ;; A typed `select` with exactly one result is the legal arity.
  (func (export "sel") (result i32)
    (select (result i32) (i32.const 4) (i32.const 5) (i32.const 1))
  )
)
(assert_return (invoke "refs") (i32.const 1))
(assert_return (invoke "mut") (i32.const 3))
(assert_return (invoke "lanes") (i32.const 16))
(assert_return (invoke "sel") (i32.const 4))

;; Extended-const arithmetic IS constant: these modules must build and run.
(module
  (memory 1)
  (data (i32.add (i32.const 0) (i32.const 42)) "x")
  (func (export "d") (result i32) (i32.load8_u (i32.const 42)))
)
(assert_return (invoke "d") (i32.const 120))

;; The other side of the positional rule: a global MAY read a global declared
;; before it, and an elem/data offset may read a defined global — all three
;; were rejected by the "imports only" reading.
(module
  (global $first i32 (i32.const 7))
  (global $second i32 (global.get $first))
  (memory 1)
  (data (global.get $first) "z")
  (table 8 funcref)
  (func $g)
  (elem (global.get $first) $g)
  (func (export "second") (result i32) (global.get $second))
  (func (export "byte") (result i32) (i32.load8_u (i32.const 7)))
)
(assert_return (invoke "second") (i32.const 7))
(assert_return (invoke "byte") (i32.const 122))

;; Tags: every spelling of a tag reference must resolve, and a tag with no
;; results is fine in all of them. `(type $t)`, inline `param`, imported, named
;; and numeric — the numeric one is the spelling that was missed.
(module
  (type $t2 (func (param i32)))
  (tag $named (param i32))
  (tag $typed (type $t2))
  (tag $none)
  (tag $multi (param i32 i64))
  ;; Each tag is THROWN somewhere, so the operand check runs against every
  ;; spelling rather than merely accepting the declarations.
  (func (export "t") (param i32) (result i32)
    (if (i32.eqz (local.get 0)) (then (throw $none)))
    (if (i32.eq (local.get 0) (i32.const 1)) (then (throw $named (i32.const 1))))
    (if (i32.eq (local.get 0) (i32.const 2)) (then (throw $typed (i32.const 2))))
    (if (i32.eq (local.get 0) (i32.const 3))
        (then (throw $multi (i32.const 3) (i64.const 4))))
    (i32.const 7)
  )
)
(assert_return (invoke "t" (i32.const 9)) (i32.const 7))

;; ── and now the rejecting half of each rule ───────────────────────────────
(assert_invalid (module (func $f (drop (ref.func $f)))) "undeclared function reference")
(assert_invalid (module (start $f) (func $f (drop (ref.func $f)))) "undeclared function reference")
(assert_invalid (module (func) (start 1)) "unknown function")
(assert_invalid (module (func $m (param i32)) (start $m)) "start function")
(assert_invalid (module (func $m (result i32) (i32.const 0)) (start $m)) "start function")
(assert_invalid (module (global f32 (f32.const 0)) (func (global.set 0 (f32.const 1)))) "immutable global")
(assert_invalid (module (global $g f32 (f32.const 0)) (func (global.set $g (f32.const 1)))) "immutable global")
(assert_invalid
  (module (import "spectest" "global_i32" (global i32)) (func (global.set 0 (i32.const 1))))
  "immutable global")
(assert_invalid (module (func (select (result) (nop) (nop) (i32.const 1)))) "invalid result arity")
(assert_invalid
  (module (func (result i32 i32)
    (select (result i32 i32) (i32.const 0) (i32.const 0) (i32.const 0) (i32.const 0) (i32.const 1))))
  "invalid result arity")
(assert_invalid
  (module (func (result i32) (i8x16.extract_lane_s 16 (v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))))
  "invalid lane index")
(assert_invalid
  (module (func (result i64) (i64x2.extract_lane 2 (v128.const i64x2 0 0))))
  "invalid lane index")
(assert_invalid (module (global i32 (global.get 0))) "unknown global")
;; ⛔ `throw 0` and `throw $x` reach the tree as DIFFERENT rules — an
;; `instr_arg` spells a bare tagidx as `integer`, not `index`, so a lookup
;; written for `id`/`index` catches the named form and silently misses the
;; numeric one. Both spellings, both forms (folded and plain), are listed.
(assert_invalid (module (func (throw 0))) "unknown tag 0")
(assert_invalid (module (func throw 0)) "unknown tag 0")
(assert_invalid (module (tag) (func (throw 1))) "unknown tag 1")
(assert_invalid (module (func (throw $missing))) "unknown tag")
(assert_invalid (module (tag (result i32))) "non-empty tag result type")
(assert_invalid (module (import "" "" (tag (result i32)))) "non-empty tag result type")
;; The suite's only two DETAILED type-mismatch wordings, both on `throw`.
(assert_invalid (module (tag (param i32)) (func (throw 0)))
  "type mismatch: instruction requires [i32] but stack has []")
(assert_invalid (module (tag (param i32)) (func (i64.const 5) (throw 0)))
  "type mismatch: instruction requires [i32] but stack has [i64]")
(assert_invalid (module (global $a i32 (global.get $b)) (global $b i32 (i32.const 0))) "unknown global")
(assert_invalid (module (global i32)) "type mismatch")
(assert_invalid (module (global i32 (i32.ctz (i32.const 0)))) "constant expression required")
(assert_invalid (module (global i32 (i32.const 0) (nop))) "constant expression required")
(assert_invalid
  (module (global (import "test" "g") (mut i32)) (memory 1) (data (global.get 0)))
  "constant expression required")
;; A TABLE initializer is allocated before globals exist, so a global declared
;; ahead of it is still out of scope — this is the fixture that proves the
;; visibility rule is positional and not one flat "imports only".
(assert_invalid
  (module (global $g funcref (ref.null func)) (table $t 10 funcref (global.get $g)))
  "unknown global")
(assert_invalid (module (global $a i32 (global.get $a))) "unknown global")

;; ── The GC / exception typing rules ─────────────────────────────────────────
;;
;; ⛔ `exnref` IS THE ENTRY FEE FOR ALL OF THESE. `parse_vt` resolved the
;; `-ref` abbreviations through the runtime's table, which lists `anyref` …
;; `nullexternref` and stops before the exception hierarchy — so every rule
;; below bailed on the SPELLING while `abs_subtype` already knew `noexn <: exn`
;; perfectly well. One absent row, an entire subsystem silently unvalidated.

;; `throw_ref` pops an exnref. Both fixtures spell it with an EMPTY stack, so
;; the pop is the whole rule — going unreachable first would eat the mismatch.
(assert_invalid (module (func (throw_ref))) "type mismatch")
(assert_invalid (module (func (block (throw_ref)))) "type mismatch")
(module (func (param exnref) (local.get 0) (throw_ref)))

;; `try_table`: the handler labels resolve in the OUTER context. This module is
;; invalid only under that reading — against the try_table's own frame, whose
;; result is empty, a `catch` carrying nothing types clean.
(assert_invalid
  (module (tag) (func (result exnref) (try_table (catch 0 0)) (unreachable)))
  "type mismatch")
(assert_invalid (module (tag) (func (try_table (catch_ref 0 0)))) "type mismatch")
(assert_invalid (module (func (try_table (catch_all_ref 0)))) "type mismatch")
;; …and a try_table is still an ORDINARY BLOCK: these two need no exception
;; machinery at all, and both went undetected while the flattener abandoned the
;; whole function on the keyword.
(assert_invalid (module (func (result i32) (try_table (result i32)))) "type mismatch")
(assert_invalid (module (func (result i32) (try_table (result i32) (i64.const 42)))) "type mismatch")
(module (tag $e (param i32)) (func (result i32) (block $h (result i32)
  (try_table (result i32) (catch $e $h) (i32.const 1)))))

;; `array.copy`: packedness is part of the storage type, and `array_elem_vt`
;; erases it — i8 and i16 both read as i32. Comparing the erased values makes
;; the first of these type clean.
(assert_invalid
  (module (type $a (array (mut i8))) (type $b (array i16))
    (func (param (ref $a)) (param (ref $b))
      (array.copy $a $b (local.get 0) (i32.const 0) (local.get 1) (i32.const 0) (i32.const 0))))
  "array types do not match")
(module (type $a (array (mut i8))) (type $b (array i8))
  (func (param (ref $a)) (param (ref $b))
    (array.copy $a $b (local.get 0) (i32.const 0) (local.get 1) (i32.const 0) (i32.const 0))))

;; `array.init_data` fills from raw bytes: a reference element has no byte form.
(assert_invalid
  (module (type $a (array (mut funcref))) (data $d "a")
    (func (param (ref $a))
      (array.init_data $a $d (local.get 0) (i32.const 0) (i32.const 0) (i32.const 0))))
  "array type is not numeric or vector")

;; `array.init_elem` resolves `$e` in the ELEMENT SEGMENT index space, not the
;; type one — looking it up among types finds nothing and abstains silently.
(assert_invalid
  (module (type $a (array (mut funcref))) (elem $e externref)
    (func (param (ref $a))
      (array.init_elem $a $e (local.get 0) (i32.const 0) (i32.const 0) (i32.const 0))))
  "type mismatch")

;; `struct.get` with a NAMED field. Field names are scoped to their struct, so
;; `$x` here is type 0's i64 — not type $t's i32 — and the declared i32 result
;; is what makes it invalid. The numeric spelling always worked; the named one
;; bailed, which is an ABSENT answer, not a wrong one.
(assert_invalid
  (module (type (struct (field $x i64))) (type $t (struct (field $x i32)))
    (func (param (ref 0)) (result i32) (struct.get 0 $x (local.get 0))))
  "type mismatch")
(module (type (struct (field $x i64))) (type $t (struct (field $x i32)))
  (func (param (ref 0)) (result i64) (struct.get 0 $x (local.get 0))))

;; `br_on_cast`: `rt2 <: rt1` is a side condition, not a consequence — casting
;; eqref to anyref WIDENS, and no stack shape reveals it.
(assert_invalid (module (func (result anyref) (br_on_cast 0 eqref anyref (unreachable))))
  "type mismatch")
(assert_invalid (module (func (result anyref) (br_on_cast_fail 0 structref arrayref (unreachable))))
  "type mismatch")
;; The fall-through carries the DIFFERENCE rt1\rt2: rt2 here is non-nullable,
;; so the null case survives and `(ref any)` cannot receive it.
(assert_invalid
  (module (type $t (struct))
    (func (param (ref null any)) (result (ref $t))
      (block (result (ref any)) (br_on_cast 1 (ref null any) (ref $t) (local.get 0))) (unreachable)))
  "type mismatch")
(module (type $t (struct))
  (func (param (ref null any)) (result (ref $t))
    (block (result (ref null any)) (br_on_cast 1 (ref null any) (ref $t) (local.get 0))) (unreachable)))

;; `table.copy` compared ADDRESS WIDTHS only; `table_of` returned the element
;; type all along and nothing read it. `table.init` names a table then a SEGMENT.
(assert_invalid
  (module (table $t1 10 funcref) (table $t2 10 externref)
    (func (table.copy $t1 $t2 (i32.const 0) (i32.const 1) (i32.const 2))))
  "type mismatch")
(assert_invalid
  (module (table $t 10 funcref) (elem $el externref)
    (func (table.init $t $el (i32.const 0) (i32.const 1) (i32.const 2))))
  "type mismatch")
(module (table $t 10 funcref) (elem $el funcref)
  (func (table.init $t $el (i32.const 0) (i32.const 1) (i32.const 2))))

;; §3.4.4: a non-nullable element type has no default, so an uninitialised
;; table cannot exist. ⛔ THE SIZE IS IRRELEVANT — `(table 0 (ref func))` is
;; invalid too; reading it as "only a non-empty table needs a default" passes
;; the 10-slot fixture and fails the 0-slot one.
(assert_invalid (module (table 0 (ref func))) "type mismatch")
(assert_invalid (module (type $f (func)) (table 10 (ref $f))) "type mismatch")
(module (table 0 funcref))
(module (table 1 (ref null func)))
(module (func $f) (table 1 (ref func) (ref.func $f)))

;; ── Local initialization (§3.4.1) ───────────────────────────────────────────
;;
;; ⛔ INITS MADE INSIDE A STRUCTURED INSTRUCTION DO NOT ESCAPE IT. The third
;; module below is the one that decides the design: `$x` is set in BOTH the
;; `then` and the `else`, and reading it after the `if` is STILL invalid. A
;; join over the branches — the intuitive reading, and the one a dataflow
;; analysis reaches for — accepts that module and fails the fixture. The frame
;; restores the init state it opened with, and `else` restores to the `if`'s.
(assert_invalid
  (module (func (local $x (ref extern)) (drop (local.get $x))))
  "uninitialized local")
(assert_invalid
  (module (func (param $p (ref extern)) (local $x (ref extern))
    (block (local.set $x (local.get $p)) (drop (local.tee $x (local.get $p))))
    (drop (local.get $x))))
  "uninitialized local")
(assert_invalid
  (module (func (param $p (ref extern)) (local $x (ref extern))
    (if (i32.const 0) (then (local.set $x (local.get $p)))
                      (else (local.set $x (local.get $p))))
    (drop (local.get $x))))
  "uninitialized local")
(assert_invalid
  (module (func (param $p (ref extern)) (local $x (ref extern))
    (if (i32.const 0) (then (local.set $x (local.get $p)))
                      (else (drop (local.get $x))))))
  "uninitialized local")
;; A DEFAULTABLE local needs no assignment, a param arrives initialized, and a
;; straight-line set is enough — all three must still compile and run.
(module (func (local $x i32) (drop (local.get $x))))
(module (func (local $x externref) (drop (local.get $x))))
(module (func (param $p (ref extern)) (drop (local.get $p))))
(module (func (param $p (ref extern)) (local $x (ref extern))
  (local.set $x (local.get $p)) (drop (local.get $x))))
;; ⛔ A `tee` WRITES BEFORE IT READS BACK, so it initializes its own local —
;; checking it as a read rejects the very instruction that makes it valid.
(module (func (param $p (ref extern)) (local $x (ref extern))
  (drop (local.tee $x (local.get $p)))))

;; §6.6.4: a numeric `(type N)` must name a type that EXISTS, counting the
;; implicit types inline signatures define — DEDUPED. Here `$g` reuses type 0
;; and `$h` reuses the implicit type 1, so the space holds 2 and `(type 2)` is
;; out of range. Counting one implicit type per inline signature (the census's
;; safe UPPER BOUND) says 4 and lets this module through.
(assert_invalid
  (module
    (func $f (result f64) (f64.const 0))
    (func $g (param i32))
    (func $h (result f64) (f64.const 1))
    (type $t (func (param i32)))
    (func (type 2)))
  "unknown type")
(module
  (func $f (result f64) (f64.const 0))
  (func $g (param i32))
  (func $h (result f64) (f64.const 1))
  (type $t (func (param i32)))
  (func (type 1) (f64.const 2)))

;; ── Iso-recursive type identity (§6.6.4) ────────────────────────────────────
;;
;; ⛔ AN INLINE SIGNATURE MAY REUSE AN EXISTING TYPE, BUT "REUSE" MEANS *THE
;; SAME TYPE* — AND IDENTITY IS ISO-RECURSIVE. A type declared inside a
;; multi-member `(rec …)` is identified by its whole group, so a standalone
;; `(func $f)` is never the same type as one of its members however exactly the
;; parameters and results line up. Recovering the function's type index by
;; SIGNATURE MATCH handed `$f` the group member's index, so the global below
;; compared a type against itself and typed clean.
;;
;; Each invalid module here differs from the valid one after it in the REC GROUP
;; ALONE — which is the only reason this pins anything.
(assert_invalid
  (module (rec (type $ft (func)) (type (func)))
    (func $f)
    (global (ref $ft) (ref.func $f)))
  "type mismatch")
(assert_invalid
  (module (rec (type $s (struct)) (type $t (func (param (ref $s)))))
    (func $f (param (ref $s)))
    (global (ref $t) (ref.func $f)))
  "type mismatch")
;; A STANDALONE type is its own singleton group, so the ordinary case must keep
;; working — this is the module the fix must not break.
(module (type $ft (func)) (func $f) (global (ref $ft) (ref.func $f)))
(module (type $s (struct)) (type $t (func (param (ref $s))))
  (func $f (param (ref $s))) (global (ref $t) (ref.func $f)))
;; ⛔ `(func $f)` DEFINES THE TYPE `[] -> []` LIKE ANY OTHER SIGNATURE. Skipping
;; the empty one left such a function with no type index, so `ref.func` bailed
;; and the whole init expression went unvalidated — absent, not conservative.
(assert_invalid
  (module (type $a (func (param i32))) (func $f) (global (ref $a) (ref.func $f)))
  "type mismatch")
