use crate::wat_exec;

wat_exec! {
    test_i64_and => { r#"
(func (export "_start")
  i64.const 0x0F
  i64.const 0x33
  i64.and
  call $log_i64
)
"#, "3" },

    test_i64_or => { r#"
(func (export "_start")
  i64.const 0x0F
  i64.const 0x30
  i64.or
  call $log_i64
)
"#, "63" },

    test_i64_xor => { r#"
(func (export "_start")
  i64.const 0xFF
  i64.const 0xAA
  i64.xor
  call $log_i64
)
"#, "85" },

    test_i64_shl => { r#"
(func (export "_start")
  i64.const 1
  i64.const 40
  i64.shl
  call $log_i64
)
"#, "1099511627776" },

    test_i64_shr_u => { r#"
(func (export "_start")
  i64.const -1
  i64.const 1
  i64.shr_u
  call $log_i64
)
"#, "9223372036854775807" },

    test_i64_shr_s => { r#"
(func (export "_start")
  i64.const -2
  i64.const 1
  i64.shr_s
  call $log_i64
)
"#, "-1" },

    test_i64_rotl => { r#"
(func (export "_start")
  i64.const -9223372036854775808 ;; 0x8000000000000000
  i64.const 1
  i64.rotl
  call $log_i64
)
"#, "1" },

    test_i64_rotr => { r#"
(func (export "_start")
  i64.const 1
  i64.const 1
  i64.rotr
  call $log_i64
)
"#, "-9223372036854775808" },

    test_i64_clz => { r#"
(func (export "_start")
  i64.const 0x0FFFFFFFFFFFFFFF
  i64.clz
  call $log_i64
)
"#, "4" },

    test_i64_clz_zero => { r#"
(func (export "_start")
  i64.const 0
  i64.clz
  call $log_i64
)
"#, "64" },

    test_i64_ctz => { r#"
(func (export "_start")
  i64.const 0x8000000000000000
  i64.ctz
  call $log_i64
)
"#, "63" },

    test_i64_ctz_zero => { r#"
(func (export "_start")
  i64.const 0
  i64.ctz
  call $log_i64
)
"#, "64" },

    test_i64_popcnt => { r#"
(func (export "_start")
  i64.const 0x0F0F0F0F0F0F0F0F
  i64.popcnt
  call $log_i64
)
"#, "32" },

    test_i64_popcnt_zero => { r#"
(func (export "_start")
  i64.const 0
  i64.popcnt
  call $log_i64
)
"#, "0" },
    
    test_i64_popcnt_all => { r#"
(func (export "_start")
  i64.const -1
  i64.popcnt
  call $log_i64
)
"#, "64" }
}
