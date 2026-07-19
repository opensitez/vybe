use crate::wat_exec;

wat_exec! {
    test_array_new_default => { r#"
(type $Arr (array i32))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 5
  array.new_default $Arr
  local.set $a
  
  local.get $a
  array.len
  call $log
)
"#, "5" },

    test_array_new => { r#"
(type $Arr (array i32))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 42
  i32.const 5
  array.new $Arr
  local.set $a
  
  local.get $a
  i32.const 0
  array.get $Arr
  call $log
)
"#, "42" },

    test_array_new_fixed => { r#"
(type $Arr (array i32))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 10
  i32.const 20
  i32.const 30
  array.new_fixed $Arr 3
  local.set $a
  
  local.get $a
  i32.const 1
  array.get $Arr
  call $log
)
"#, "20" },

    // A fixed GC array is stamped with its rtt, so an out-of-bounds access
    // traps per WASM spec (unlike a lenient dynamic array).
    test_array_new_fixed_get_oob => { r#"
(type $Arr (array i32))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 10
  i32.const 20
  i32.const 30
  array.new_fixed $Arr 3
  local.set $a

  local.get $a
  i32.const 9
  array.get $Arr
  call $log
)
"#, "trap" },

    test_array_new_data => { r#"
(type $Arr (array (mut i8)))
(data $d "data")
(func (export "_start") (local $a (ref null $Arr))
  i32.const 0
  i32.const 4
  array.new_data $Arr $d
  local.set $a
  
  local.get $a
  i32.const 0
  array.get_u $Arr
  call $log
)
"#, "100" }, // 'd' = 100

    test_array_new_elem => { r#"
(type $Arr (array (mut funcref)))
(func $f)
(elem $e $f)
(func (export "_start") (local $a (ref null $Arr))
  i32.const 0
  i32.const 1
  array.new_elem $Arr $e
  local.set $a

  local.get $a
  array.len
  call $log
)
"#, "1" }
}
