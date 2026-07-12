use crate::wat_exec;

wat_exec! {
    test_array_subtype_get => { r#"
(type $Base (array (mut i32)))
(type $Sub (array_subtype (mut i32) $Base))
(func (export "_start") (local $a (ref null $Base))
  i32.const 42
  i32.const 5
  array.new $Sub
  local.set $a
  
  local.get $a
  i32.const 2
  array.get $Base
  call $log
)
"#, "42" },

    test_array_subtype_set => { r#"
(type $Base (array (mut i32)))
(type $Sub (array_subtype (mut i32) $Base))
(func (export "_start") (local $a (ref null $Base))
  i32.const 10
  i32.const 5
  array.new $Sub
  local.set $a
  
  local.get $a
  i32.const 2
  i32.const 99
  array.set $Base
  
  local.get $a
  i32.const 2
  array.get $Base
  call $log
)
"#, "99" },

    test_array_subtype_func_param => { r#"
(type $Base (array (mut i32)))
(type $Sub (array_subtype (mut i32) $Base))
(func $f1 (param $a (ref null $Base)) (result i32)
  local.get $a
  i32.const 0
  array.get $Base)
(func (export "_start")
  i32.const 99
  i32.const 5
  array.new $Sub
  call $f1
  call $log
)
"#, "99" },

    test_array_subtype_func_return => { r#"
(type $Base (array (mut i32)))
(type $Sub (array_subtype (mut i32) $Base))
(func $f1 (result (ref null $Base))
  i32.const 42
  i32.const 5
  array.new $Sub)
(func (export "_start")
  call $f1
  i32.const 0
  array.get $Base
  call $log
)
"#, "42" }
}
