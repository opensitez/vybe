use crate::wat_exec;

wat_exec! {
    test_ref_cast_success => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  i32.const 20
  struct.new $Sub
  local.set $s
  
  local.get $s
  ref.cast $Sub
  struct.get $Sub 1
  call $log
)
"#, "20" },

    test_ref_cast_fail => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  struct.new $Base
  local.set $s
  
  local.get $s
  ref.cast $Sub
  drop
  i32.const 42
  call $log
)
"#, "trap" },

    test_ref_cast_null => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  ref.null $Base
  local.set $s
  
  local.get $s
  ref.cast $Sub
  drop
  i32.const 42
  call $log
)
"#, "trap" },

    test_ref_cast_null_success => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  ref.null $Base
  local.set $s
  
  local.get $s
  ref.cast (ref null $Sub)
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_ref_test_success => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  i32.const 20
  struct.new $Sub
  local.set $s
  
  local.get $s
  ref.test $Sub
  call $log
)
"#, "1" },

    test_ref_test_fail => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  struct.new $Base
  local.set $s
  
  local.get $s
  ref.test $Sub
  call $log
)
"#, "0" },

    test_ref_test_null => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  ref.null $Base
  local.set $s
  
  local.get $s
  ref.test $Sub
  call $log
)
"#, "0" },

    test_br_on_cast_success => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  i32.const 20
  struct.new $Sub
  local.set $s
  
  block (result (ref null $Sub))
    local.get $s
    br_on_cast 0 $Base $Sub
    drop
    i32.const 0
    i32.const 0
    struct.new $Sub
  end
  struct.get $Sub 1
  call $log
)
"#, "20" },

    test_br_on_cast_fail => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  struct.new $Base
  local.set $s
  
  block (result (ref null $Sub))
    local.get $s
    br_on_cast 0 $Base $Sub
    drop
    i32.const 99
    i32.const 88
    struct.new $Sub
  end
  struct.get $Sub 1
  call $log
)
"#, "88" },

    test_br_on_cast_fail_null => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  ref.null $Base
  local.set $s
  
  block (result (ref null $Sub))
    local.get $s
    br_on_cast 0 $Base $Sub
    drop
    i32.const 99
    i32.const 88
    struct.new $Sub
  end
  struct.get $Sub 1
  call $log
)
"#, "88" }
}
