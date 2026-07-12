use crate::wat_exec;

wat_exec! {
    test_memory_load_i32 => { r#"
(memory 1)
(data (i32.const 0) "\01\02\03\04")
(func (export "_start") i32.const 0 i32.load call $log)
"#, "67305985" }, // 0x04030201

    test_memory_load_i64 => { r#"
(memory 1)
(data (i32.const 8) "\01\02\03\04\05\06\07\08")
(func (export "_start") i32.const 8 i64.load call $log_i64)
"#, "578437695752307201" }, // 0x0807060504030201

    test_memory_load_f32 => { r#"
(memory 1)
(data (i32.const 0) "\00\00\80\3f") ;; 1.0 in f32
(func (export "_start") i32.const 0 f32.load call $log_f32)
"#, "1.0" },

    test_memory_load_f64 => { r#"
(memory 1)
(data (i32.const 0) "\00\00\00\00\00\00\f0\3f") ;; 1.0 in f64
(func (export "_start") i32.const 0 f64.load call $log_f64)
"#, "1.0" },

    test_memory_load_offset => { r#"
(memory 1)
(data (i32.const 10) "\10\20\30\40")
(func (export "_start") i32.const 5 i32.load offset=5 call $log)
"#, "1076895760" }, // 0x40302010

    test_memory_load_align_1 => { r#"
(memory 1)
(data (i32.const 1) "\10\20\30\40")
(func (export "_start") i32.const 1 i32.load align=1 call $log)
"#, "1076895760" }, // 0x40302010

    test_memory_load8_s_pos => { r#"
(memory 1)
(data (i32.const 0) "\7f")
(func (export "_start") i32.const 0 i32.load8_s call $log)
"#, "127" },

    test_memory_load8_s_neg => { r#"
(memory 1)
(data (i32.const 0) "\80")
(func (export "_start") i32.const 0 i32.load8_s call $log)
"#, "-128" },

    test_memory_load8_u => { r#"
(memory 1)
(data (i32.const 0) "\80")
(func (export "_start") i32.const 0 i32.load8_u call $log)
"#, "128" },

    test_memory_load16_s_pos => { r#"
(memory 1)
(data (i32.const 0) "\ff\7f")
(func (export "_start") i32.const 0 i32.load16_s call $log)
"#, "32767" },

    test_memory_load16_s_neg => { r#"
(memory 1)
(data (i32.const 0) "\00\80")
(func (export "_start") i32.const 0 i32.load16_s call $log)
"#, "-32768" },

    test_memory_load16_u => { r#"
(memory 1)
(data (i32.const 0) "\00\80")
(func (export "_start") i32.const 0 i32.load16_u call $log)
"#, "32768" },

    test_memory_load_i64_8_s_pos => { r#"
(memory 1)
(data (i32.const 0) "\7f")
(func (export "_start") i32.const 0 i64.load8_s call $log_i64)
"#, "127" },

    test_memory_load_i64_8_s_neg => { r#"
(memory 1)
(data (i32.const 0) "\80")
(func (export "_start") i32.const 0 i64.load8_s call $log_i64)
"#, "-128" },

    test_memory_load_i64_8_u => { r#"
(memory 1)
(data (i32.const 0) "\80")
(func (export "_start") i32.const 0 i64.load8_u call $log_i64)
"#, "128" },

    test_memory_load_i64_16_s_neg => { r#"
(memory 1)
(data (i32.const 0) "\00\80")
(func (export "_start") i32.const 0 i64.load16_s call $log_i64)
"#, "-32768" },

    test_memory_load_i64_16_u => { r#"
(memory 1)
(data (i32.const 0) "\00\80")
(func (export "_start") i32.const 0 i64.load16_u call $log_i64)
"#, "32768" },

    test_memory_load_i64_32_s_pos => { r#"
(memory 1)
(data (i32.const 0) "\ff\ff\ff\7f")
(func (export "_start") i32.const 0 i64.load32_s call $log_i64)
"#, "2147483647" },

    test_memory_load_i64_32_s_neg => { r#"
(memory 1)
(data (i32.const 0) "\00\00\00\80")
(func (export "_start") i32.const 0 i64.load32_s call $log_i64)
"#, "-2147483648" },

    test_memory_load_i64_32_u => { r#"
(memory 1)
(data (i32.const 0) "\00\00\00\80")
(func (export "_start") i32.const 0 i64.load32_u call $log_i64)
"#, "2147483648" },

    test_memory_load_oob_trap => { r#"
(memory 1)
(func (export "_start") i32.const 65536 i32.load call $log)
"#, "trap" },

    test_memory_load_partially_oob_trap => { r#"
(memory 1)
(func (export "_start") i32.const 65533 i32.load call $log)
"#, "trap" }
}
