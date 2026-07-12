use crate::wat_exec;

wat_exec! {
    test_elem_segment_active => { r#"
(table 5 funcref)
(func $f1 (result i32) i32.const 42)
(elem (i32.const 0) $f1)
(func (export "_start")
  i32.const 0
  table.get 0
  ref.is_null
  call $log
)
"#, "0" },

    test_elem_segment_passive_init => { r#"
(table 5 funcref)
(func $f1 (result i32) i32.const 42)
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
"#, "0" },

    test_elem_segment_active_overlap => { r#"
(table 5 funcref)
(func $f1)
(func $f2)
(elem (i32.const 0) $f1)
(elem (i32.const 0) $f2)
(func (export "_start")
  i32.const 0
  table.get 0
  ref.func $f2
  ref.eq
  call $log
)
"#, "1" },

    test_elem_segment_out_of_bounds => { r#"
(table 2 funcref)
(func $f1)
(elem (i32.const 5) $f1)
(func (export "_start")
  i32.const 42
  call $log
)
"#, "trap" }, // fails to instantiate

    test_elem_segment_declarative => { r#"
(func $f1 (result i32) i32.const 42)
(elem declare $f1)
(func (export "_start")
  ref.func $f1
  drop
  i32.const 42
  call $log
)
"#, "42" }, // allows ref.func without putting it in a table
    
    test_elem_drop => { r#"
(table 5 funcref)
(func $f1 (result i32) i32.const 42)
(elem $e $f1)
(func (export "_start")
  elem.drop $e
  i32.const 0
  i32.const 0
  i32.const 1
  table.init $e
  i32.const 0
  call $log
)
"#, "trap" } // dropped elem segment -> oob
}
