use crate::wat_exec;

wat_exec! {
    test_v128_const => { r#"
(func (export "_start")
  v128.const i32x4 0x01020304 0x05060708 0x090A0B0C 0x0D0E0F10
  i32x4.extract_lane 0
  call $log
)
"#, "16909060" },

    test_v128_load => { r#"
(memory 1)
(data (i32.const 0) "\04\03\02\01\08\07\06\05\0c\0b\0a\09\10\0f\0e\0d")
(func (export "_start")
  i32.const 0
  v128.load
  i32x4.extract_lane 3
  call $log
)
"#, "219025168" }, // 0x0d0e0f10

    test_v128_store => { r#"
(memory 1)
(func (export "_start")
  i32.const 0
  v128.const i32x4 42 99 100 200
  v128.store
  i32.const 4
  i32.load
  call $log
)
"#, "99" },

    test_v128_load8x8_s => { r#"
(memory 1)
(data (i32.const 0) "\ff\00\80\7f\01\02\03\04")
(func (export "_start")
  i32.const 0
  v128.load8x8_s
  i16x8.extract_lane_s 2
  call $log
)
"#, "-128" },

    test_v128_load8x8_u => { r#"
(memory 1)
(data (i32.const 0) "\ff\00\80\7f\01\02\03\04")
(func (export "_start")
  i32.const 0
  v128.load8x8_u
  i16x8.extract_lane_s 2
  call $log
)
"#, "128" },

    test_v128_load16x4_s => { r#"
(memory 1)
(data (i32.const 0) "\ff\ff\00\00\00\80\ff\7f")
(func (export "_start")
  i32.const 0
  v128.load16x4_s
  i32x4.extract_lane 2
  call $log
)
"#, "-32768" },

    test_v128_load16x4_u => { r#"
(memory 1)
(data (i32.const 0) "\ff\ff\00\00\00\80\ff\7f")
(func (export "_start")
  i32.const 0
  v128.load16x4_u
  i32x4.extract_lane 2
  call $log
)
"#, "32768" },

    test_v128_load32x2_s => { r#"
(memory 1)
(data (i32.const 0) "\ff\ff\ff\ff\00\00\00\80")
(func (export "_start")
  i32.const 0
  v128.load32x2_s
  i64x2.extract_lane 1
  call $log_i64
)
"#, "-2147483648" },

    test_v128_load32x2_u => { r#"
(memory 1)
(data (i32.const 0) "\ff\ff\ff\ff\00\00\00\80")
(func (export "_start")
  i32.const 0
  v128.load32x2_u
  i64x2.extract_lane 1
  call $log_i64
)
"#, "2147483648" },

    test_v128_load32_zero => { r#"
(memory 1)
(data (i32.const 0) "\ff\ff\ff\ff")
(func (export "_start")
  i32.const 0
  v128.load32_zero
  i32x4.extract_lane 1
  call $log
)
"#, "0" }, // higher lanes are zeroed

    test_v128_load64_zero => { r#"
(memory 1)
(data (i32.const 0) "\ff\ff\ff\ff\ff\ff\ff\ff")
(func (export "_start")
  i32.const 0
  v128.load64_zero
  i64x2.extract_lane 1
  call $log_i64
)
"#, "0" } // higher lane is zeroed
}
