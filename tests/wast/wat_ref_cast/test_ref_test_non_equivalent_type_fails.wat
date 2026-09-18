;; vybe-test: wast/wat_ref_cast/test_ref_test_non_equivalent_type_fails
;; origin: languages/wast/tests/wast/test_wat_ref_cast.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  ;; `Step_read/ref.test` answers `$Ref_ok(ref)` matched against the immediate,
  ;; and Defined Types match only when EQUAL CLOSED under the context. $A and
  ;; $C are not equivalent — different field storage — and neither is a subtype
  ;; of the other, so the answer is 0 in both directions.
  ;;
  ;; ⛔ $A and $B WOULD be equivalent: matching is ISO-RECURSIVE
  ;; (`clostype_C(deftype_1) = clostype_C(deftype_2)`), so two separately
  ;; declared types with an identical shape are ONE type and `ref.test` answers
  ;; 1 for either. Asserting 0 there pins a nominal rule the spec does not have.
  (type $A (struct (field i32)))
(type $C (struct (field i64)))
(func (export "_start") (local $a (ref null $A)) (local $c (ref null $C))
  i32.const 10
  struct.new $A
  local.set $a
  i64.const 20
  struct.new $C
  local.set $c
  
  local.get $a
  ref.test $A
  i32.const 1 call $vybe_check_i32
  
  local.get $a
  ref.test $C
  i32.const 0 call $vybe_check_i32
  
  local.get $c
  ref.test $A
  i32.const 0 call $vybe_check_i32
)
)
