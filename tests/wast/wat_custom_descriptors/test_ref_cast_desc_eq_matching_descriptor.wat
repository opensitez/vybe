;; vybe-test: wast/wat_custom_descriptors/test_ref_cast_desc_eq_matching_descriptor
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/ref_cast_desc_eq.wast

;; The cast compares the reference's descriptor against the supplied one by
;; IDENTITY ("descriptor equality", i.e. the same reference). When they match
;; the reference passes through unchanged, exactly like `ref.cast`.

(module
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)

  (rec
    (type $a (descriptor $a.desc) (struct (field i32)))
    (type $a.desc (describes $a) (struct (field i32)))
  )

  (func (export "_start") (local $d (ref null $a.desc)) (local $o (ref null $a))
    i32.const 5
    struct.new $a.desc
    local.set $d

    i32.const 1
    local.get $d
    struct.new_desc $a
    local.set $o

    ;; Same descriptor value ⇒ the cast succeeds and yields the reference.
    local.get $o
    local.get $d
    ref.cast_desc_eq (ref $a)
    struct.get $a 0
    i32.const 1
    call $vybe_check_i32
  )
)
