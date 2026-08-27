;; vybe-test: wast/wat_spec_gc/type_immediates_may_be_positional
;; vybe-test-mode: run
;;
;; A type immediate may be written POSITIONALLY — `struct.get_s 0 0` names
;; type 0 — exactly as a func, table or memory index may. Every per-type map
;; in the walker is keyed by the declared NAME, and the numeric spelling was
;; being passed through as the literal string `"0"`, so every one of those
;; lookups missed.
;;
;; The visible consequence is a WRONG ANSWER, not an error: the packed field's
;; storage type was unknown, so `struct.get_s` skipped its sign extension and
;; answered 254 where the spec says -2. `gc/struct.wast` writes its whole
;; packed-field section positionally, and `get_u` agreed with `get_s` on every
;; value — which is precisely the symptom, since both were reading raw.

(module
  ;; Unnamed, so the only way to address it is by index.
  (type (struct (field i8) (field (mut i8)) (field i16) (field (mut i16))))

  (global (ref 0) (struct.new 0 (i32.const 254) (i32.const 255)
                                (i32.const 65534) (i32.const 65535)))

  (func (export "i8_s") (result i32) (struct.get_s 0 0 (global.get 0)))
  (func (export "i8_u") (result i32) (struct.get_u 0 0 (global.get 0)))
  (func (export "i16_s") (result i32) (struct.get_s 0 2 (global.get 0)))
  (func (export "i16_u") (result i32) (struct.get_u 0 2 (global.get 0)))

  ;; A positional `struct.new_default` must find the same field types, so the
  ;; defaults are the storage type's and not a uniform null.
  (func (export "default_i8_u") (result i32)
    (struct.get_u 0 0 (struct.new_default 0)))

  ;; A positional `struct.set` addresses the same slot a positional get does.
  (func (export "set_then_get") (result i32)
    (local $v (ref null 0))
    (local.set $v (struct.new_default 0))
    (struct.set 0 1 (local.get $v) (i32.const 200))
    (struct.get_s 0 1 (local.get $v))))

(assert_return (invoke "i8_s") (i32.const -2))
(assert_return (invoke "i8_u") (i32.const 254))
(assert_return (invoke "i16_s") (i32.const -2))
(assert_return (invoke "i16_u") (i32.const 65534))
(assert_return (invoke "default_i8_u") (i32.const 0))
(assert_return (invoke "set_then_get") (i32.const -56))

;; ── The NAMED spelling must keep answering the same ─────────────────
;; A resolver that mapped names through the index table too would break these.
(module
  (type $s (struct (field i8) (field $b (mut i8))))
  (global $g (ref $s) (struct.new $s (i32.const 254) (i32.const 255)))
  (func (export "n_s") (result i32) (struct.get_s $s 0 (global.get $g)))
  (func (export "n_u") (result i32) (struct.get_u $s 0 (global.get $g)))
  (func (export "n_by_name") (result i32) (struct.get_u $s $b (global.get $g))))

(assert_return (invoke "n_s") (i32.const -2))
(assert_return (invoke "n_u") (i32.const 254))
(assert_return (invoke "n_by_name") (i32.const 255))
