;; Native Wasm class/type gate, independent of the PHP source compiler.
;; A shared descriptor carries a typed method reference; two descriptors of
;; the same structural type still have distinct runtime reference identities.
(module
  (rec
    (type $instance (descriptor $descriptor)
      (struct (field $value (mut i32))))
    (type $descriptor (describes $instance)
      (struct (field $class i32) (field $method (ref $method))))
    (type $method (func (param (ref null $instance)) (result i32)))
  )
  (elem declare func $read)
  (func $read (type $method) (param $receiver (ref null $instance)) (result i32)
    (struct.get $instance $value (local.get $receiver)))
  (global $first (ref (exact $descriptor))
    (struct.new $descriptor (i32.const 1) (ref.func $read)))
  (global $second (ref (exact $descriptor))
    (struct.new $descriptor (i32.const 1) (ref.func $read)))

  (func $check
    (local $receiver (ref null $instance))
    (local.set $receiver
      (struct.new_desc $instance (i32.const 7) (global.get $first)))
    (if (i32.eqz (ref.test (ref $instance) (local.get $receiver)))
      (then unreachable))
    (if (i32.eqz
      (ref.eq (ref.get_desc $instance (local.get $receiver)) (global.get $first)))
      (then unreachable))
    (if
      (ref.eq (ref.get_desc $instance (local.get $receiver)) (global.get $second))
      (then unreachable))
    ;; Method dispatch through the descriptor's immutable typed funcref slot.
    (if (i32.ne
      (call_ref $method
        (local.get $receiver)
        (struct.get $descriptor $method
          (ref.get_desc $instance (local.get $receiver))))
      (i32.const 7))
      (then unreachable))
    (if (i32.ne
      (struct.get $instance $value
        (ref.cast_desc_eq (ref $instance)
          (local.get $receiver) (global.get $first)))
      (i32.const 7))
      (then unreachable))
  )
  ;; Checks run during instantiation, not only if a runner chooses an export.
  (start $check)
  (func (export "main") (result i32) (i32.const 1))
)
