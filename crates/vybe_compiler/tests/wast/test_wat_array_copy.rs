use crate::wat_exec;

wat_exec! {
    test_array_copy => { r#"
(type $Arr (array (mut i32)))
(func (export "_start") (local $a1 (ref null $Arr)) (local $a2 (ref null $Arr))
  i32.const 10
  i32.const 5
  array.new $Arr
  local.set $a1
  
  i32.const 20
  i32.const 5
  array.new $Arr
  local.set $a2
  
  local.get $a2
  i32.const 1
  local.get $a1
  i32.const 0
  i32.const 3
  array.copy $Arr $Arr
  
  local.get $a2
  i32.const 1
  array.get $Arr
  call $log
)
"#, "10" },

    test_array_copy_overlap => { r#"
(type $Arr (array (mut i32)))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 10
  i32.const 20
  i32.const 30
  array.new_fixed $Arr 3
  local.set $a
  
  local.get $a
  i32.const 1
  local.get $a
  i32.const 0
  i32.const 2
  array.copy $Arr $Arr
  
  local.get $a
  i32.const 1
  array.get $Arr
  call $log
)
"#, "10" },

    test_array_copy_oob_dest => { r#"
(type $Arr (array (mut i32)))
(func (export "_start") (local $a1 (ref null $Arr)) (local $a2 (ref null $Arr))
  i32.const 10
  i32.const 5
  array.new $Arr
  local.set $a1
  
  i32.const 20
  i32.const 5
  array.new $Arr
  local.set $a2
  
  local.get $a2
  i32.const 4
  local.get $a1
  i32.const 0
  i32.const 3
  array.copy $Arr $Arr
  
  i32.const 0
  call $log
)
"#, "trap" },

    test_array_copy_oob_src => { r#"
(type $Arr (array (mut i32)))
(func (export "_start") (local $a1 (ref null $Arr)) (local $a2 (ref null $Arr))
  i32.const 10
  i32.const 5
  array.new $Arr
  local.set $a1
  
  i32.const 20
  i32.const 5
  array.new $Arr
  local.set $a2
  
  local.get $a2
  i32.const 0
  local.get $a1
  i32.const 4
  i32.const 3
  array.copy $Arr $Arr
  
  i32.const 0
  call $log
)
"#, "trap" },

    test_array_copy_null_dest => { r#"
(type $Arr (array (mut i32)))
(func (export "_start") (local $a1 (ref null $Arr)) (local $a2 (ref null $Arr))
  i32.const 10
  i32.const 5
  array.new $Arr
  local.set $a1
  
  ref.null $Arr
  local.set $a2
  
  local.get $a2
  i32.const 0
  local.get $a1
  i32.const 0
  i32.const 3
  array.copy $Arr $Arr
  
  i32.const 0
  call $log
)
"#, "trap" },

    test_array_copy_null_src => { r#"
(type $Arr (array (mut i32)))
(func (export "_start") (local $a1 (ref null $Arr)) (local $a2 (ref null $Arr))
  ref.null $Arr
  local.set $a1
  
  i32.const 20
  i32.const 5
  array.new $Arr
  local.set $a2
  
  local.get $a2
  i32.const 0
  local.get $a1
  i32.const 0
  i32.const 3
  array.copy $Arr $Arr
  
  i32.const 0
  call $log
)
"#, "trap" }
}
