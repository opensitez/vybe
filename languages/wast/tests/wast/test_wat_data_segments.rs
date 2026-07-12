use crate::wat_exec;

wat_exec! {
    test_data_segment_active => { r#"
(memory 1)
(data (i32.const 10) "hello")
(func (export "_start")
  i32.const 10
  i32.load8_u
  call $log
)
"#, "104" }, // 'h'

    test_data_segment_passive_init => { r#"
(memory 1)
(data $d "hello")
(func (export "_start")
  i32.const 10
  i32.const 0
  i32.const 5
  memory.init $d
  i32.const 10
  i32.load8_u
  call $log
)
"#, "104" },

    test_data_segment_active_overlap => { r#"
(memory 1)
(data (i32.const 0) "hello")
(data (i32.const 1) "world")
(func (export "_start")
  i32.const 1
  i32.load8_u
  call $log
)
"#, "119" }, // 'w'

    test_data_segment_out_of_bounds => { r#"
(memory 1)
(data (i32.const 65535) "hello")
(func (export "_start")
  i32.const 42
  call $log
)
"#, "trap" }, // fails to instantiate

    test_data_segment_multi_memory => { r#"
(memory $m1 1)
(memory $m2 1)
(data (memory $m2) (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.load8_u $m2
  call $log
)
"#, "104" }
}
