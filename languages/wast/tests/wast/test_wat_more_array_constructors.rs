//! GC "more array constructors" — array.new, new_default, new_fixed, new_data,
//! new_elem, plus array.len and element access.
use crate::wat_exec;

wat_exec! {
    test_array_new_fixed => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start")
          i32.const 10 i32.const 20 i32.const 30 array.new_fixed $a 3
          i32.const 1 array.get $a call $log))"#, "20" },
    test_array_new_with_value_and_len => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start")
          i32.const 7 i32.const 5 array.new $a i32.const 3 array.get $a call $log))"#, "7" },
    test_array_new_default_is_zero => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start")
          i32.const 4 array.new_default $a i32.const 0 array.get $a call $log))"#, "0" },
    test_array_len => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start")
          i32.const 0 i32.const 6 array.new $a array.len call $log))"#, "6" },
    test_array_new_data => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (data $d "\63\00\00\00\64\00\00\00")
        (func (export "_start")
          i32.const 0 i32.const 2 array.new_data $a $d i32.const 0 array.get $a call $log))"#, "99" },
    test_array_set_then_get => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start") (local $arr (ref $a))
          i32.const 0 i32.const 5 array.new $a local.set $arr
          local.get $arr i32.const 2 i32.const 42 array.set $a
          local.get $arr i32.const 2 array.get $a call $log))"#, "42" },
}
