//! GC ops not covered elsewhere — i31 references (ref.i31, i31.get_s/u),
//! array.fill, array.init_data / array.init_elem, and any/extern conversions.
use crate::wat_exec;

wat_exec! {
    test_ref_i31_and_get_s => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") i32.const 42 ref.i31 i31.get_s call $log))"#, "42" },
    test_i31_get_s_sign_extends => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") i32.const -1 ref.i31 i31.get_s call $log))"#, "-1" },
    test_i31_get_u_zero_extends => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") i32.const -1 ref.i31 i31.get_u call $log))"#, "2147483647" },
    test_i31_truncates_to_31_bits => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") i32.const 0x7FFFFFFF ref.i31 i31.get_u call $log))"#, "2147483647" },
    test_array_fill => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start") (local $arr (ref $a))
          i32.const 0 i32.const 5 array.new $a local.set $arr
          local.get $arr i32.const 1 i32.const 9 i32.const 3 array.fill $a
          local.get $arr i32.const 2 array.get $a call $log))"#, "9" },
    test_array_init_data => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (data $d "\07\00\00\00\08\00\00\00")
        (func (export "_start") (local $arr (ref $a))
          i32.const 0 i32.const 4 array.new_default $a local.set $arr
          local.get $arr i32.const 0 i32.const 0 i32.const 2 array.init_data $a $d
          local.get $arr i32.const 1 array.get $a call $log))"#, "8" },
    test_array_copy => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (type $a (array (mut i32)))
        (func (export "_start") (local $src (ref $a)) (local $dst (ref $a))
          i32.const 55 i32.const 3 array.new $a local.set $src
          i32.const 0 i32.const 3 array.new_default $a local.set $dst
          local.get $dst i32.const 0 local.get $src i32.const 0 i32.const 3 array.copy $a $a
          local.get $dst i32.const 1 array.get $a call $log))"#, "55" },
    test_extern_any_convert_roundtrip => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          i32.const 42 ref.i31 extern.convert_any any.convert_extern
          ref.cast (ref i31) i31.get_s call $log))"#, "42" },
}

