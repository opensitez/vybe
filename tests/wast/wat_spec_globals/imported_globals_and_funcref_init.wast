;; vybe-test: wast/wat_spec_globals/imported_globals_and_funcref_init
;; vybe-test-mode: run
;;
;; Two spec behaviours that were broken and are now fixed; this file is what
;; found both and what guards them.
;;
;; 1. IMPORTED GLOBALS WERE NOT LINKED. `walk_global_field` never read
;;    `import_inline`, so `(global (import "m" "g") i32)` became an ordinary
;;    local global left at its `Expression::int(0)` default — every read of an
;;    imported global answered 0. Imported FUNCTIONS and MEMORIES already
;;    worked, which is what scoped this to globals rather than to linking.
;;    The fix aliases the importer's binding to the exporter's, so an imported
;;    MUTABLE global is one shared cell rather than a copy — asserted below in
;;    both directions.
;;
;; 2. A `ref.func` GLOBAL INITIALISER read back NULL. `ref.func $f` lowers to
;;    `<ModuleClass>.$f` and globals were declared BEFORE the class, so the
;;    initialiser read an undefined class. Such globals are now declared after
;;    it (still before `start`).

(module $M
  (global (export "imm") i32 (i32.const 5))
  (global (export "mut") (mut i32) (i32.const 7))
  (global (export "f64imm") f64 (f64.const 1.5))
  (func (export "bump") (global.set 1 (i32.add (global.get 1) (i32.const 1))))
  (func (export "read_mut") (result i32) (global.get 1))
)
(register "m")

(module
  (global $imm (import "m" "imm") i32)
  (global $mut (import "m" "mut") (mut i32))
  (global $f64imm (import "m" "f64imm") f64)

  ;; The ONE `global.get` a constant expression may use: an imported,
  ;; immutable global.
  (global $derived i32 (global.get $imm))
  (global $derived_f64 f64 (global.get $f64imm))

  (elem declare func $ten)
  (func $ten (result i32) (i32.const 10))
  (global $funcref_init funcref (ref.func $ten))
  (table 1 funcref)
  (type $thunk (func (result i32)))

  (func (export "derived") (result i32) (global.get $derived))
  (func (export "derived_f64") (result f64) (global.get $derived_f64))
  (func (export "read_imported_imm") (result i32) (global.get $imm))
  (func (export "read_imported_mut") (result i32) (global.get $mut))
  (func (export "write_imported_mut") (param i32) (global.set $mut (local.get 0)))
  (func (export "funcref_is_null") (result i32) (ref.is_null (global.get $funcref_init)))
  (func (export "call_from_global") (result i32)
    (table.set (i32.const 0) (global.get $funcref_init))
    (call_indirect (type $thunk) (i32.const 0)))
)

;; ── (1) imported globals ───────────────────────────────────────────────
(assert_return (invoke "read_imported_imm") (i32.const 5))
(assert_return (invoke "read_imported_mut") (i32.const 7))
;; An initialiser reading an imported global takes ITS value.
(assert_return (invoke "derived") (i32.const 5))
(assert_return (invoke "derived_f64") (f64.const 1.5))
;; An imported mutable global is ONE cell shared with the exporter, not a copy.
(invoke "write_imported_mut" (i32.const 55))
(assert_return (invoke "read_imported_mut") (i32.const 55))
(assert_return (invoke $M "read_mut") (i32.const 55))
(invoke $M "bump")
(assert_return (invoke $M "read_mut") (i32.const 56))
(assert_return (invoke "read_imported_mut") (i32.const 56))
;; The derived global is a SNAPSHOT of the init value — mutating the exporter's
;; mutable global does not disturb a global initialised from the immutable one.
(assert_return (invoke "derived") (i32.const 5))

;; ── (2) ref.func initialiser ───────────────────────────────────────────
(assert_return (invoke "funcref_is_null") (i32.const 0))
(assert_return (invoke "call_from_global") (i32.const 10))
