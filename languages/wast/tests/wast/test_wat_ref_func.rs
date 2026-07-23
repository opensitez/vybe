use crate::wat_exec;

wat_exec! {
    test_ref_func => { r#"
(func $f1 (result i32) i32.const 42)
(func (export "_start") (local $r funcref)
  ref.func $f1
  local.set $r
  local.get $r
  ref.is_null
  call $log
)
"#, "0" },

    test_ref_func_call_indirect => { r#"
(type $sig (func (result i32)))
(table 1 funcref)
(func $f1 (result i32) i32.const 42)
(func (export "_start")
  i32.const 0
  ref.func $f1
  table.set 0
  
  i32.const 0
  call_indirect (type $sig)
  call $log
)
"#, "42" },

    test_ref_func_call_ref => { r#"
(type $sig (func (result i32)))
(func $f1 (result i32) i32.const 42)
(func (export "_start") (local $r (ref null $sig))
  ref.func $f1
  local.set $r
  
  local.get $r
  call_ref $sig
  call $log
)
"#, "42" },

    test_call_ref_null => { r#"
(type $sig (func (result i32)))
(func (export "_start") (local $r (ref null $sig))
  ref.null $sig
  local.set $r
  
  local.get $r
  call_ref $sig
  call $log
)
"#, "trap" },

    test_call_ref_args => { r#"
(type $sig (func (param i32 i32) (result i32)))
(func $add (param i32 i32) (result i32)
  local.get 0
  local.get 1
  i32.add)
(func (export "_start") (local $r (ref null $sig))
  ref.func $add
  local.set $r
  
  i32.const 10
  i32.const 20
  local.get $r
  call_ref $sig
  call $log
)
"#, "30" },

    test_return_call_ref => { r#"
(type $sig (func (result i32)))
(func $f1 (result i32) i32.const 42)
(func (export "_start") (result i32) (local $r (ref null $sig))
  ref.func $f1
  local.set $r

  local.get $r
  return_call_ref $sig
)
"#, "42" },

    test_return_call_ref_args => { r#"
(type $sig (func (param i32 i32) (result i32)))
(func $add (param i32 i32) (result i32)
  local.get 0
  local.get 1
  i32.add)
(func (export "_start") (result i32) (local $r (ref null $sig))
  ref.func $add
  local.set $r

  i32.const 10
  i32.const 20
  local.get $r
  return_call_ref $sig
)
"#, "30" },

    test_return_call_ref_null => { r#"
(type $sig (func (result i32)))
(func (export "_start") (local $r (ref null $sig))
  ref.null $sig
  local.set $r
  
  local.get $r
  return_call_ref $sig
)
"#, "trap" }
}
