use crate::wat_exec;

wat_exec! {
    test_call_indirect_success => { r#"
(type $sig (func (result i32)))
(table 2 funcref)
(func $f1 (result i32) i32.const 42)
(func $f2 (result i32) i32.const 99)
(elem (i32.const 0) $f1 $f2)
(func (export "_start")
  i32.const 0
  call_indirect (type $sig)
  call $log
)
"#, "42" },

    test_call_indirect_second => { r#"
(type $sig (func (result i32)))
(table 2 funcref)
(func $f1 (result i32) i32.const 42)
(func $f2 (result i32) i32.const 99)
(elem (i32.const 0) $f1 $f2)
(func (export "_start")
  i32.const 1
  call_indirect (type $sig)
  call $log
)
"#, "99" },

    test_call_indirect_args => { r#"
(type $sig (func (param i32 i32) (result i32)))
(table 1 funcref)
(func $add (param i32 i32) (result i32)
  local.get 0
  local.get 1
  i32.add)
(elem (i32.const 0) $add)
(func (export "_start")
  i32.const 10
  i32.const 20
  i32.const 0
  call_indirect (type $sig)
  call $log
)
"#, "30" },

    test_call_indirect_oob => { r#"
(type $sig (func (result i32)))
(table 2 funcref)
(func $f1 (result i32) i32.const 42)
(func $f2 (result i32) i32.const 99)
(elem (i32.const 0) $f1 $f2)
(func (export "_start")
  i32.const 2
  call_indirect (type $sig)
  call $log
)
"#, "trap" },

    test_call_indirect_null => { r#"
(type $sig (func (result i32)))
(table 2 funcref)
(func $f1 (result i32) i32.const 42)
(elem (i32.const 0) $f1)
;; index 1 is null
(func (export "_start")
  i32.const 1
  call_indirect (type $sig)
  call $log
)
"#, "trap" },

    test_call_indirect_signature_mismatch_params => { r#"
(type $sig1 (func (result i32)))
(type $sig2 (func (param i32) (result i32)))
(table 1 funcref)
(func $f1 (type $sig2) 
  local.get 0)
(elem (i32.const 0) $f1)
(func (export "_start")
  i32.const 0
  call_indirect (type $sig1) ;; calling a func that takes 1 param as if it takes 0
  call $log
)
"#, "trap" },

    test_call_indirect_signature_mismatch_results => { r#"
(type $sig1 (func))
(type $sig2 (func (result i32)))
(table 1 funcref)
(func $f1 (type $sig2) 
  i32.const 42)
(elem (i32.const 0) $f1)
(func (export "_start")
  i32.const 0
  call_indirect (type $sig1) ;; calling a func that returns 1 result as if it returns 0
  i32.const 0
  call $log
)
"#, "trap" },

    test_call_indirect_multiple_tables => { r#"
(type $sig (func (result i32)))
(table $t1 1 funcref)
(table $t2 1 funcref)
(func $f1 (result i32) i32.const 42)
(func $f2 (result i32) i32.const 99)
(elem (table $t1) (i32.const 0) $f1)
(elem (table $t2) (i32.const 0) $f2)
(func (export "_start")
  i32.const 0
  call_indirect $t2 (type $sig)
  call $log
)
"#, "99" }
}
