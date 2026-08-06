//! Typed function references proposal — ref.func, call_ref, ref.as_non_null,
//! br_on_null / br_on_non_null with typed function references.
use crate::wat_exec;

wat_exec! {
    test_call_ref_invokes_function => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $t (func (result i32)))
        (func $answer (type $t) i32.const 42)
        (func (export "_start") ref.func $answer call_ref $t call $log))"#, "42" },
    test_call_ref_with_args => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $bin (func (param i32 i32) (result i32)))
        (func $add (type $bin) local.get 0 local.get 1 i32.add)
        (func (export "_start")
          i32.const 20 i32.const 22 ref.func $add call_ref $bin call $log))"#, "42" },
    test_ref_as_non_null => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $t (func (result i32)))
        (func $f (type $t) i32.const 7)
        (func (export "_start") ref.func $f ref.as_non_null call_ref $t call $log))"#, "7" },
    test_br_on_null_not_taken => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $t (func (result i32)))
        (func $f (type $t) i32.const 5)
        (func (export "_start") (result i32)
          block (result (ref $t))
            ref.func $f br_on_null 0
            call_ref $t return
          end drop i32.const -1)
        (func (export "_run") call $log))"#, "5" },
    test_function_ref_stored_and_called => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $t (func (result i32)))
        (func $a (type $t) i32.const 100)
        (func $b (type $t) i32.const 200)
        (func $pick (param $which i32) (result (ref $t))
          local.get $which if (result (ref $t)) ref.func $a else ref.func $b end)
        (func (export "_start") i32.const 0 call $pick call_ref $t call $log))"#, "200" },
}
