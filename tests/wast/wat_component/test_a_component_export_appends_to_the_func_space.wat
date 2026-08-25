;; vybe-test: wast/wat_component/test_a_component_export_appends_to_the_func_space
;; hand-written against proposals/component-model/design/mvp/Explainer.md:2601
;;
;; > not only import definitions, but also export definitions append a new
;; > element to the index space of the imported/exported `sort` … In the case
;; > of exports, the `<id>?` right after the `export` is bound while the
;; > `<id>` inside the `<externidx>` is a reference to the preceding definition
;; > being exported (e.g., `(export $x "x" (func $f))` binds a new identifier
;; > `$x`).
;;
;; ▶▶ THE TWO `$id`s IN AN EXPORT ARE OPPOSITES. `$f` READS the function space;
;; the leading `$x` NAMES THE NEW ENTRY the export adds to it. Reading the
;; leading one as a second name for `$f` looks correct until something indexes
;; positionally past the export — from there every index is off by one, which
;; is the `GLOBAL_GET` shape: one integer meaning two things depending on who
;; reads it.
;;
;; The component function index space here is:
;;
;;   funcidx 0 → `canon lift` of `times2`
;;   funcidx 1 → `canon lift` of `times3`
;;   funcidx 2 → the EXPORT of funcidx 1
;;
;; ⛔ TWO DISCRIMINATORS, BOTH REQUIRED:
;;
;; 1. `canon lower` names funcidx **2 POSITIONALLY**. A `$id`-only test cannot
;;    tell an appending export from a renaming one — both bind the id to
;;    something callable. If the export only rebound a name, funcidx 2 would
;;    not exist and this refuses with `component func 2 is not defined
;;    (have 2)`. The companion check is that `(func 3)` refuses with `have 3`,
;;    which pins the space to exactly three entries.
;; 2. There are **TWO** lifts and the export names the SECOND. With one lifted
;;    function an export resolving to the wrong index would land back on the
;;    same function and pass anyway. 14 × 3 = 42; resolving to funcidx 0 gives
;;    28 instead.

(component
  (core module $m
    (func (export "times2") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 2)))
    (func (export "times3") (param i32) (result i32)
      (i32.mul (local.get 0) (i32.const 3))))
  (core instance $mi (instantiate $m))
  (alias core export $mi "times2" (core func $c2))
  (alias core export $mi "times3" (core func $c3))

  (type $ft (func (param "a" u32) (result u32)))
  (canon lift (core func $c2) (func $times2 (type $ft)))
  (canon lift (core func $c3) (func $times3 (type $ft)))

  (export "run" (func $times3))

  (canon lower (func 2) (core func $lo))

  (core module $caller
    (import "canon" "lo" (func $l (param i32) (result i32)))
    (func (export "get") (result i32)
      (call $l (i32.const 14))))
  (core instance (instantiate $caller
    (with "canon" (instance (export "lo" (func $lo))))))
)

(assert_return (invoke "get") (i32.const 42))
