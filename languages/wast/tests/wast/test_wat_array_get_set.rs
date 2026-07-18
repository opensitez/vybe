use crate::wat_exec;

wat_exec! {
    // array.new_default of a `(ref null $t)` element type fills typed nulls, so
    // accessing through a defaulted ref element traps per spec.
    test_array_new_default_ref_elem_traps => { r#"
(type $Inner (struct (field i32)))
(type $A (array (ref null $Inner)))
(func (export "_start")
  i32.const 2
  array.new_default $A
  i32.const 0
  array.get $A
  struct.get $Inner 0
  call $log
)
"#, "trap" },

    test_array_get => { r#"
(type $Arr (array i32))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 42
  i32.const 5
  array.new $Arr
  local.set $a
  
  local.get $a
  i32.const 2
  array.get $Arr
  call $log
)
"#, "42" },

    test_array_get_oob => { r#"
(type $Arr (array i32))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 42
  i32.const 5
  array.new $Arr
  local.set $a
  
  local.get $a
  i32.const 5
  array.get $Arr
  call $log
)
"#, "trap" },

    test_array_get_null => { r#"
(type $Arr (array i32))
(func (export "_start") (local $a (ref null $Arr))
  ref.null $Arr
  local.set $a
  
  local.get $a
  i32.const 0
  array.get $Arr
  call $log
)
"#, "trap" },

    test_array_get_s => { r#"
(type $Arr (array i8))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 255 ;; -1 as i8
  i32.const 1
  array.new $Arr
  local.set $a
  
  local.get $a
  i32.const 0
  array.get_s $Arr
  call $log
)
"#, "-1" },

    test_array_get_u => { r#"
(type $Arr (array i8))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 255
  i32.const 1
  array.new $Arr
  local.set $a
  
  local.get $a
  i32.const 0
  array.get_u $Arr
  call $log
)
"#, "255" },

    test_array_set => { r#"
(type $Arr (array (mut i32)))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 0
  i32.const 5
  array.new $Arr
  local.set $a
  
  local.get $a
  i32.const 2
  i32.const 42
  array.set $Arr
  
  local.get $a
  i32.const 2
  array.get $Arr
  call $log
)
"#, "42" },

    test_array_set_oob => { r#"
(type $Arr (array (mut i32)))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 0
  i32.const 5
  array.new $Arr
  local.set $a
  
  local.get $a
  i32.const 5
  i32.const 42
  array.set $Arr
  
  i32.const 0
  call $log
)
"#, "trap" },

    test_array_set_null => { r#"
(type $Arr (array (mut i32)))
(func (export "_start") (local $a (ref null $Arr))
  ref.null $Arr
  local.set $a
  
  local.get $a
  i32.const 0
  i32.const 42
  array.set $Arr
  
  i32.const 0
  call $log
)
"#, "trap" }
}
