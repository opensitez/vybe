use crate::wat_exec;

wat_exec! {
    test_table_get => { r#"
(table 2 funcref)
(func $f1 (result i32) i32.const 42)
(elem (i32.const 0) $f1)
(func (export "_start")
  i32.const 0
  table.get 0
  ref.as_non_null
  drop
  i32.const 1
  call $log
)
"#, "1" },

    test_table_get_null => { r#"
(table 2 funcref)
(func $f1 (result i32) i32.const 42)
(elem (i32.const 0) $f1)
(func (export "_start")
  i32.const 1
  table.get 0
  ref.is_null
  call $log
)
"#, "1" },

    test_table_get_oob => { r#"
(table 2 funcref)
(func (export "_start")
  i32.const 2
  table.get 0
  drop
  i32.const 1
  call $log
)
"#, "trap" },

    test_table_set => { r#"
(table 2 funcref)
(func $f1)
(func (export "_start")
  i32.const 1
  ref.func $f1
  table.set 0
  i32.const 1
  table.get 0
  ref.is_null
  call $log
)
"#, "0" },

    test_table_set_oob => { r#"
(table 2 funcref)
(func $f1)
(func (export "_start")
  i32.const 2
  ref.func $f1
  table.set 0
  i32.const 1
  call $log
)
"#, "trap" },

    test_table_size => { r#"
(table 5 funcref)
(func (export "_start")
  table.size 0
  call $log
)
"#, "5" },

    test_table_grow_success => { r#"
(table 5 funcref)
(func (export "_start")
  ref.null func
  i32.const 2
  table.grow 0
  call $log
)
"#, "5" }, ;; returns old size

    test_table_grow_fail => { r#"
(table 5 5 funcref)
(func (export "_start")
  ref.null func
  i32.const 1
  table.grow 0
  call $log
)
"#, "-1" },

    test_table_size_after_grow => { r#"
(table 5 funcref)
(func (export "_start")
  ref.null func
  i32.const 3
  table.grow 0
  drop
  table.size 0
  call $log
)
"#, "8" },

    test_table_fill => { r#"
(table 5 funcref)
(func $f1)
(func (export "_start")
  i32.const 1
  ref.func $f1
  i32.const 3
  table.fill 0
  i32.const 2
  table.get 0
  ref.is_null
  call $log
)
"#, "0" },

    test_table_fill_oob => { r#"
(table 5 funcref)
(func $f1)
(func (export "_start")
  i32.const 3
  ref.func $f1
  i32.const 3
  table.fill 0
  i32.const 1
  call $log
)
"#, "trap" },

    test_table_copy => { r#"
(table $t1 5 funcref)
(table $t2 5 funcref)
(func $f1)
(elem (table $t1) (i32.const 2) $f1)
(func (export "_start")
  i32.const 0
  i32.const 2
  i32.const 1
  table.copy $t2 $t1
  i32.const 0
  table.get $t2
  ref.is_null
  call $log
)
"#, "0" },
    
    test_table_init => { r#"
(table 5 funcref)
(func $f1)
(elem $e $f1)
(func (export "_start")
  i32.const 1
  i32.const 0
  i32.const 1
  table.init $e
  i32.const 1
  table.get 0
  ref.is_null
  call $log
)
"#, "0" }
}
