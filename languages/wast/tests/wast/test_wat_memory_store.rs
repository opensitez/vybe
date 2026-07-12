use crate::wat_exec;

wat_exec! {
    test_memory_store_i32 => { r#"
(memory 1)
(func (export "_start") 
  i32.const 0 
  i32.const 16909060 ;; 0x01020304
  i32.store 
  i32.const 0 
  i32.load 
  call $log)
"#, "16909060" },

    test_memory_store_i64 => { r#"
(memory 1)
(func (export "_start") 
  i32.const 8 
  i64.const 578437695752307201 ;; 0x0807060504030201
  i64.store 
  i32.const 8 
  i64.load 
  call $log_i64)
"#, "578437695752307201" },

    test_memory_store_f32 => { r#"
(memory 1)
(func (export "_start") 
  i32.const 0 
  f32.const 1.0 
  f32.store 
  i32.const 0 
  f32.load 
  call $log_f32)
"#, "1.0" },

    test_memory_store_f64 => { r#"
(memory 1)
(func (export "_start") 
  i32.const 0 
  f64.const 1.0 
  f64.store 
  i32.const 0 
  f64.load 
  call $log_f64)
"#, "1.0" },

    test_memory_store_offset => { r#"
(memory 1)
(func (export "_start") 
  i32.const 5 
  i32.const 42
  i32.store offset=10 
  i32.const 15 
  i32.load 
  call $log)
"#, "42" },

    test_memory_store_align => { r#"
(memory 1)
(func (export "_start") 
  i32.const 1 
  i32.const 42
  i32.store align=1 
  i32.const 1 
  i32.load 
  call $log)
"#, "42" },

    test_memory_store8_i32 => { r#"
(memory 1)
(func (export "_start") 
  i32.const 0 
  i32.const 300 ;; 256 + 44
  i32.store8 
  i32.const 0 
  i32.load8_u 
  call $log)
"#, "44" },

    test_memory_store16_i32 => { r#"
(memory 1)
(func (export "_start") 
  i32.const 0 
  i32.const 65580 ;; 65536 + 44
  i32.store16 
  i32.const 0 
  i32.load16_u 
  call $log)
"#, "44" },

    test_memory_store8_i64 => { r#"
(memory 1)
(func (export "_start") 
  i32.const 0 
  i64.const 300 
  i64.store8 
  i32.const 0 
  i64.load8_u 
  call $log_i64)
"#, "44" },

    test_memory_store16_i64 => { r#"
(memory 1)
(func (export "_start") 
  i32.const 0 
  i64.const 65580 
  i64.store16 
  i32.const 0 
  i64.load16_u 
  call $log_i64)
"#, "44" },

    test_memory_store32_i64 => { r#"
(memory 1)
(func (export "_start") 
  i32.const 0 
  i64.const 4294967340 ;; 4294967296 + 44
  i64.store32 
  i32.const 0 
  i64.load32_u 
  call $log_i64)
"#, "44" },

    test_memory_store_oob_trap => { r#"
(memory 1)
(func (export "_start") 
  i32.const 65536 
  i32.const 1 
  i32.store 
  i32.const 0
  call $log)
"#, "trap" },

    test_memory_store_partially_oob_trap => { r#"
(memory 1)
(func (export "_start") 
  i32.const 65533 
  i32.const 1 
  i32.store 
  i32.const 0
  call $log)
"#, "trap" }
}
