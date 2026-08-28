;; vybe-test: wast/wat_custom_descriptors/test_struct_new_default_desc
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/struct_new_desc.wast

;; `struct.new_default_desc $a` takes ONLY the descriptor: every field is set to
;; its storage type's default. The descriptor still round-trips.

(module
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)

  (rec
    (type $a (descriptor $a.desc) (struct (field i32) (field i32)))
    (type $a.desc (describes $a) (struct (field i32)))
  )

  (func (export "_start") (local $d (ref null $a.desc)) (local $o (ref null $a))
    i32.const 99
    struct.new $a.desc
    local.set $d

    local.get $d
    struct.new_default_desc $a
    local.set $o

    ;; Both fields defaulted.
    local.get $o
    struct.get $a 0
    i32.const 0
    call $vybe_check_i32
    local.get $o
    struct.get $a 1
    i32.const 0
    call $vybe_check_i32

    ;; The descriptor is the one that was supplied.
    local.get $o
    ref.get_desc $a
    struct.get $a.desc 0
    i32.const 99
    call $vybe_check_i32
  )
)
