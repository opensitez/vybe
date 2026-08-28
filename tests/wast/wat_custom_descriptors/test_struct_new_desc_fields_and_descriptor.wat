;; vybe-test: wast/wat_custom_descriptors/test_struct_new_desc_fields_and_descriptor
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/struct_new_desc.wast

;; `struct.new_desc $a` takes the field values FIRST and the descriptor LAST
;; (Overview.md §"Allocation With Descriptors"). The descriptor is on top of
;; the stack, above the fields.
;;
;; This is the case the previous implementation got wrong: it popped only the
;; descriptor and ignored both the type index and the field operands, so the
;; field value below it was left stranded on the stack and the allocation came
;; back as an empty untyped object.

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)

  ;; The described and describing types refer to each other, so they must
  ;; share a recursion group.
  (rec
    (type $a (descriptor $a.desc) (struct (field i32)))
    (type $a.desc (describes $a) (struct (field i32)))
  )

  (func (export "_start") (local $d (ref null $a.desc)) (local $o (ref null $a))
    ;; A descriptor type has a `describes` clause but no `descriptor` clause,
    ;; so it is allocated with plain `struct.new` — no chicken-and-egg.
    i32.const 7
    struct.new $a.desc
    local.set $d

    i32.const 42
    local.get $d
    struct.new_desc $a
    local.set $o

    ;; The field operand landed in the instance rather than being stranded.
    local.get $o
    struct.get $a 0
    i32.const 42
    call $vybe_check_i32

    ;; ...and the descriptor is retrievable with `ref.get_desc`.
    local.get $o
    ref.get_desc $a
    struct.get $a.desc 0
    i32.const 7
    call $vybe_check_i32
  )
)
