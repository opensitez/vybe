use crate::wat_exec;

wat_exec! {
    test_i32_trunc_f32_s_pos => { r#"
(func (export "_start")
  f32.const 42.9
  i32.trunc_f32_s
  call $log
)
"#, "42" },

    test_i32_trunc_f32_s_neg => { r#"
(func (export "_start")
  f32.const -42.9
  i32.trunc_f32_s
  call $log
)
"#, "-42" },

    test_i32_trunc_f32_u_pos => { r#"
(func (export "_start")
  f32.const 42.9
  i32.trunc_f32_u
  call $log
)
"#, "42" },

    test_i32_trunc_f32_u_trap_neg => { r#"
(func (export "_start")
  f32.const -1.0
  i32.trunc_f32_u
  call $log
)
"#, "trap" },

    test_i64_trunc_f64_s_pos => { r#"
(func (export "_start")
  f64.const 42.9
  i64.trunc_f64_s
  call $log_i64
)
"#, "42" },

    test_i64_trunc_f64_s_neg => { r#"
(func (export "_start")
  f64.const -42.9
  i64.trunc_f64_s
  call $log_i64
)
"#, "-42" },

    test_f32_convert_i32_s_pos => { r#"
(func (export "_start")
  i32.const 42
  f32.convert_i32_s
  call $log_f32
)
"#, "42.0" },

    test_f32_convert_i32_s_neg => { r#"
(func (export "_start")
  i32.const -42
  f32.convert_i32_s
  call $log_f32
)
"#, "-42.0" },

    test_f64_convert_i64_u_pos => { r#"
(func (export "_start")
  i64.const 42
  f64.convert_i64_u
  call $log_f64
)
"#, "42.0" },

    test_f64_promote_f32 => { r#"
(func (export "_start")
  f32.const 42.5
  f64.promote_f32
  call $log_f64
)
"#, "42.5" },

    test_f32_demote_f64 => { r#"
(func (export "_start")
  f64.const 42.5
  f32.demote_f64
  call $log_f32
)
"#, "42.5" },

    test_f32_reinterpret_i32 => { r#"
(func (export "_start")
  i32.const 1065353216 ;; 0x3f800000 = 1.0f
  f32.reinterpret_i32
  call $log_f32
)
"#, "1.0" },

    test_i32_reinterpret_f32 => { r#"
(func (export "_start")
  f32.const 1.0
  i32.reinterpret_f32
  call $log
)
"#, "1065353216" },

    test_f64_reinterpret_i64 => { r#"
(func (export "_start")
  i64.const 4607182418800017408 ;; 0x3FF0000000000000 = 1.0
  f64.reinterpret_i64
  call $log_f64
)
"#, "1.0" },

    test_i64_reinterpret_f64 => { r#"
(func (export "_start")
  f64.const 1.0
  i64.reinterpret_f64
  call $log_i64
)
"#, "4607182418800017408" },

    test_i32_trunc_sat_f32_s_pos => { r#"
(func (export "_start")
  f32.const 3000000000.0 ;; > i32.max
  i32.trunc_sat_f32_s
  call $log
)
"#, "2147483647" },

    test_i32_trunc_sat_f32_s_neg => { r#"
(func (export "_start")
  f32.const -3000000000.0 ;; < i32.min
  i32.trunc_sat_f32_s
  call $log
)
"#, "-2147483648" },

    test_i32_trunc_sat_f32_s_nan => { r#"
(func (export "_start")
  f32.const nan
  i32.trunc_sat_f32_s
  call $log
)
"#, "0" }
}
