;; vybe-test: wast/wat_spec_gc/struct_fields_by_name_and_the_field_abbreviation
;; vybe-test-mode: run
;;
;; `struct.get $T $y` addresses a field BY NAME. The lowering only accepted an
;; integer literal there and fell back to index **0** for anything else, so a
;; named field always read and wrote slot 0.
;;
;; That is invisible to the obvious test. A `set $y` / `get $y` pair both go to
;; slot 0 and round-trip perfectly; so does a `set 1` / `get 1` pair. Only
;; MIXING the two spellings on one field shows it, which is exactly what
;; `gc/struct.wast`'s `set_get_1` does — and it is the only assertion in that
;; file that did.
;;
;; The second half here is the `(field t1 t2 t3)` abbreviation: it declares
;; THREE consecutive unnamed fields, not one. Reading only the first storage
;; type made every later field's index too high by two.

(module
  (type $vec (struct (field f32) (field $y (mut f32)) (field $z f32)))

  ;; Set by INDEX, read by NAME — the mixed pair.
  (func $set_1_get_y (param $v (ref $vec)) (param $val f32) (result f32)
    (struct.set $vec 1 (local.get $v) (local.get $val))
    (struct.get $vec $y (local.get $v)))
  (func (export "set_1_get_y") (param $val f32) (result f32)
    (call $set_1_get_y (struct.new_default $vec) (local.get $val)))

  ;; Set by NAME, read by INDEX — the other direction.
  (func $set_y_get_1 (param $v (ref $vec)) (param $val f32) (result f32)
    (struct.set $vec $y (local.get $v) (local.get $val))
    (struct.get $vec 1 (local.get $v)))
  (func (export "set_y_get_1") (param $val f32) (result f32)
    (call $set_y_get_1 (struct.new_default $vec) (local.get $val)))

  ;; Writing the NAMED field must leave field 0 alone. With the old fallback
  ;; both went to slot 0, so this read answered the written value.
  (func $write_y_read_0 (param $v (ref $vec)) (result f32)
    (struct.set $vec $y (local.get $v) (f32.const 9))
    (struct.get $vec 0 (local.get $v)))
  (func (export "write_y_read_0") (result f32)
    (call $write_y_read_0 (struct.new $vec (f32.const 1) (f32.const 2) (f32.const 3))))

  ;; A constructed struct read back by name and by index must agree.
  (func $read (param $v (ref $vec)) (result f32) (struct.get $vec $z (local.get $v)))
  (func (export "read_z_by_name") (result f32)
    (call $read (struct.new $vec (f32.const 1) (f32.const 2) (f32.const 3))))
)

(assert_return (invoke "set_1_get_y" (f32.const 7)) (f32.const 7))
(assert_return (invoke "set_y_get_1" (f32.const 7)) (f32.const 7))
(assert_return (invoke "write_y_read_0") (f32.const 1))
(assert_return (invoke "read_z_by_name") (f32.const 3))

;; ── The `(field t1 t2 …)` abbreviation ──────────────────────────────
;; `(field i32 i32 i32)` is THREE fields. If only the first were counted,
;; `$b`'s index would be 1 instead of 3 and these would read each other's
;; values.
(module
  (type $t (struct (field i32 i32 i32) (field $b i64) (field $c i32)))
  (func $new (result (ref $t))
    (struct.new $t (i32.const 10) (i32.const 20) (i32.const 30)
                   (i64.const 40) (i32.const 50)))
  (func (export "f0") (result i32) (struct.get $t 0 (call $new)))
  (func (export "f2") (result i32) (struct.get $t 2 (call $new)))
  (func (export "b_by_index") (result i64) (struct.get $t 3 (call $new)))
  (func (export "b_by_name") (result i64) (struct.get $t $b (call $new)))
  (func (export "c_by_index") (result i32) (struct.get $t 4 (call $new)))
  (func (export "c_by_name") (result i32) (struct.get $t $c (call $new))))

(assert_return (invoke "f0") (i32.const 10))
(assert_return (invoke "f2") (i32.const 30))
(assert_return (invoke "b_by_index") (i64.const 40))
(assert_return (invoke "b_by_name") (i64.const 40))
(assert_return (invoke "c_by_index") (i32.const 50))
(assert_return (invoke "c_by_name") (i32.const 50))
