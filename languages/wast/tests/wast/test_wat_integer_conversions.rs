use crate::wat_exec;

wat_exec! {
    test_i32_wrap_i64 => { r#"
(func (export "_start")
  i64.const 4294967338 ;; 0x10000002A -> wraps to 42
  i32.wrap_i64
  call $log
)
"#, "42" },

    test_i64_extend_i32_s_pos => { r#"
(func (export "_start")
  i32.const 42
  i64.extend_i32_s
  call $log_i64
)
"#, "42" },

    test_i64_extend_i32_s_neg => { r#"
(func (export "_start")
  i32.const -1
  i64.extend_i32_s
  call $log_i64
)
"#, "-1" },

    test_i64_extend_i32_u_pos => { r#"
(func (export "_start")
  i32.const 42
  i64.extend_i32_u
  call $log_i64
)
"#, "42" },

    test_i64_extend_i32_u_neg => { r#"
(func (export "_start")
  i32.const -1
  i64.extend_i32_u
  call $log_i64
)
"#, "4294967295" }, // 0xFFFFFFFF

    test_i32_extend8_s_pos => { r#"
(func (export "_start")
  i32.const 127
  i32.extend8_s
  call $log
)
"#, "127" },

    test_i32_extend8_s_neg => { r#"
(func (export "_start")
  i32.const 255 ;; -1 as i8
  i32.extend8_s
  call $log
)
"#, "-1" },

    test_i32_extend16_s_pos => { r#"
(func (export "_start")
  i32.const 32767
  i32.extend16_s
  call $log
)
"#, "32767" },

    test_i32_extend16_s_neg => { r#"
(func (export "_start")
  i32.const 65535 ;; -1 as i16
  i32.extend16_s
  call $log
)
"#, "-1" },

    test_i64_extend8_s_pos => { r#"
(func (export "_start")
  i64.const 127
  i64.extend8_s
  call $log_i64
)
"#, "127" },

    test_i64_extend8_s_neg => { r#"
(func (export "_start")
  i64.const 255
  i64.extend8_s
  call $log_i64
)
"#, "-1" },

    test_i64_extend16_s_pos => { r#"
(func (export "_start")
  i64.const 32767
  i64.extend16_s
  call $log_i64
)
"#, "32767" },

    test_i64_extend16_s_neg => { r#"
(func (export "_start")
  i64.const 65535
  i64.extend16_s
  call $log_i64
)
"#, "-1" },

    test_i64_extend32_s_pos => { r#"
(func (export "_start")
  i64.const 2147483647
  i64.extend32_s
  call $log_i64
)
"#, "2147483647" },

    test_i64_extend32_s_neg => { r#"
(func (export "_start")
  i64.const 4294967295 ;; -1 as i32
  i64.extend32_s
  call $log_i64
)
"#, "-1" }
}
