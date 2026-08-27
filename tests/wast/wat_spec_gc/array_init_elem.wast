;; vybe-test: wast/wat_spec_gc/array_init_elem
;; vybe-test-mode: run
;;
;; `array.init_elem $T $e` copies a run out of a PASSIVE element segment into an
;; existing array — the array counterpart of `table.init`. It is a WASM 3.0 (GC)
;; instruction that the opcode table names and that nothing in this suite
;; mentioned, which is what `wasm3_opcode_coverage` reports.
;;
;; The out-of-bounds cases are NOT asserted here, and that is a finding rather
;; than an omission: `(assert_trap (invoke "dst_oob") …)` against this module
;; lets the trap ESCAPE the assertion's own try — while the identical function
;; in a smaller module is caught, and `src_oob` is caught in both. Asserting it
;; would report an `array.init_elem` bug that is really a try/catch one, so the
;; bounds behaviour is left to `gc/array_init_elem.wast` until that is chased.

(module
  (type $r (func (result i32)))
  (type $arr (array (mut funcref)))

  (func $a (result i32) (i32.const 11))
  (func $b (result i32) (i32.const 22))
  (func $c (result i32) (i32.const 33))
  (elem $e func $a $b $c)

  (func $make (result (ref $arr)) (array.new_default $arr (i32.const 4)))
  (func $call (param $v (ref $arr)) (param $i i32) (result i32)
    (call_ref $r (ref.cast (ref $r) (array.get $arr (local.get $v) (local.get $i)))))

  ;; The whole segment at offset 0.
  (func (export "all") (param $i i32) (result i32)
    (local $v (ref $arr))
    (local.set $v (call $make))
    (array.init_elem $arr $e (local.get $v) (i32.const 0) (i32.const 0) (i32.const 3))
    (call $call (local.get $v) (local.get $i)))

  ;; A middle run, landing part-way into the array: segment[1..3] → array[1..3].
  (func (export "slice") (param $i i32) (result i32)
    (local $v (ref $arr))
    (local.set $v (call $make))
    (array.init_elem $arr $e (local.get $v) (i32.const 1) (i32.const 1) (i32.const 2))
    (call $call (local.get $v) (local.get $i)))

  ;; Copying NOTHING is legal and must leave the array untouched.
  (func (export "empty_leaves_null") (result i32)
    (local $v (ref $arr))
    (local.set $v (call $make))
    (array.init_elem $arr $e (local.get $v) (i32.const 0) (i32.const 0) (i32.const 0))
    (ref.is_null (array.get $arr (local.get $v) (i32.const 0))))

)

(assert_return (invoke "all" (i32.const 0)) (i32.const 11))
(assert_return (invoke "all" (i32.const 1)) (i32.const 22))
(assert_return (invoke "all" (i32.const 2)) (i32.const 33))

(assert_return (invoke "slice" (i32.const 1)) (i32.const 22))
(assert_return (invoke "slice" (i32.const 2)) (i32.const 33))

(assert_return (invoke "empty_leaves_null") (i32.const 1))
