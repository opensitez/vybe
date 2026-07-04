use crate::wat_exec;

wat_exec! {
    test_memory_copy => { r#"
(memory 1)
(data (i32.const 10) "hello")
(func (export "_start")
  i32.const 20 
  i32.const 10 
  i32.const 5
  memory.copy
  i32.const 20
  i32.load8_u
  call $log)
"#, "104" }, // 'h' = 104

    test_memory_copy_overlap_forward => { r#"
(memory 1)
(data (i32.const 10) "abcdef")
(func (export "_start")
  i32.const 12 
  i32.const 10 
  i32.const 4
  memory.copy
  i32.const 12
  i32.load8_u
  call $log)
"#, "97" }, // 'a' = 97

    test_memory_copy_overlap_backward => { r#"
(memory 1)
(data (i32.const 10) "abcdef")
(func (export "_start")
  i32.const 8 
  i32.const 10 
  i32.const 4
  memory.copy
  i32.const 8
  i32.load8_u
  call $log)
"#, "97" }, // 'a'

    test_memory_copy_oob_dest => { r#"
(memory 1)
(func (export "_start")
  i32.const 65530
  i32.const 0
  i32.const 10
  memory.copy
  i32.const 0
  call $log)
"#, "trap" },

    test_memory_copy_oob_src => { r#"
(memory 1)
(func (export "_start")
  i32.const 0
  i32.const 65530
  i32.const 10
  memory.copy
  i32.const 0
  call $log)
"#, "trap" },

    test_memory_copy_zero_length => { r#"
(memory 1)
(func (export "_start")
  i32.const 65536
  i32.const 0
  i32.const 0
  memory.copy
  i32.const 0
  call $log)
"#, "0" }, // OOB pointer is allowed if length is 0

    test_memory_fill => { r#"
(memory 1)
(func (export "_start")
  i32.const 10
  i32.const 255
  i32.const 5
  memory.fill
  i32.const 12
  i32.load8_u
  call $log)
"#, "255" },

    test_memory_fill_trunc => { r#"
(memory 1)
(func (export "_start")
  i32.const 10
  i32.const 257 ;; wraps to 1
  i32.const 5
  memory.fill
  i32.const 12
  i32.load8_u
  call $log)
"#, "1" },

    test_memory_fill_oob => { r#"
(memory 1)
(func (export "_start")
  i32.const 65530
  i32.const 1
  i32.const 10
  memory.fill
  i32.const 0
  call $log)
"#, "trap" },

    test_memory_fill_zero_length => { r#"
(memory 1)
(func (export "_start")
  i32.const 65536
  i32.const 1
  i32.const 0
  memory.fill
  i32.const 0
  call $log)
"#, "0" },

    test_memory_init => { r#"
(memory 1)
(data $d "data")
(func (export "_start")
  i32.const 10
  i32.const 0
  i32.const 4
  memory.init $d
  i32.const 10
  i32.load8_u
  call $log)
"#, "100" }, // 'd' = 100

    test_memory_init_partial => { r#"
(memory 1)
(data $d "data")
(func (export "_start")
  i32.const 10
  i32.const 1
  i32.const 2
  memory.init $d
  i32.const 10
  i32.load8_u
  call $log)
"#, "97" }, // 'a' = 97

    test_memory_init_oob_dest => { r#"
(memory 1)
(data $d "data")
(func (export "_start")
  i32.const 65534
  i32.const 0
  i32.const 4
  memory.init $d
  i32.const 0
  call $log)
"#, "trap" },

    test_memory_init_oob_src => { r#"
(memory 1)
(data $d "data")
(func (export "_start")
  i32.const 10
  i32.const 2
  i32.const 4
  memory.init $d
  i32.const 0
  call $log)
"#, "trap" },

    test_data_drop => { r#"
(memory 1)
(data $d "data")
(func (export "_start")
  data.drop $d
  i32.const 10
  i32.const 0
  i32.const 4
  memory.init $d
  i32.const 0
  call $log)
"#, "trap" } // dropped data segment -> oob
}
