use crate::wat_exec;

wat_exec! {
    test_i32_and => { r#"
(func (export "_start")
  i32.const 0x0F
  i32.const 0x33
  i32.and
  call $log
)
"#, "3" },

    test_i32_or => { r#"
(func (export "_start")
  i32.const 0x0F
  i32.const 0x30
  i32.or
  call $log
)
"#, "63" }, // 0x3F

    test_i32_xor => { r#"
(func (export "_start")
  i32.const 0xFF
  i32.const 0xAA
  i32.xor
  call $log
)
"#, "85" }, // 0x55

    test_i32_shl => { r#"
(func (export "_start")
  i32.const 1
  i32.const 3
  i32.shl
  call $log
)
"#, "8" },

    test_i32_shr_u => { r#"
(func (export "_start")
  i32.const -1 ;; 0xFFFFFFFF
  i32.const 1
  i32.shr_u
  call $log
)
"#, "2147483647" }, // 0x7FFFFFFF

    test_i32_shr_s => { r#"
(func (export "_start")
  i32.const -2
  i32.const 1
  i32.shr_s
  call $log
)
"#, "-1" },

    test_i32_rotl => { r#"
(func (export "_start")
  i32.const -2147483648 ;; 0x80000000
  i32.const 1
  i32.rotl
  call $log
)
"#, "1" },

    test_i32_rotr => { r#"
(func (export "_start")
  i32.const 1
  i32.const 1
  i32.rotr
  call $log
)
"#, "-2147483648" },

    test_i32_clz => { r#"
(func (export "_start")
  i32.const 0x0FFFFFFF
  i32.clz
  call $log
)
"#, "4" },

    test_i32_clz_zero => { r#"
(func (export "_start")
  i32.const 0
  i32.clz
  call $log
)
"#, "32" },

    test_i32_ctz => { r#"
(func (export "_start")
  i32.const 0x80000000
  i32.ctz
  call $log
)
"#, "31" },

    test_i32_ctz_zero => { r#"
(func (export "_start")
  i32.const 0
  i32.ctz
  call $log
)
"#, "32" },

    test_i32_popcnt => { r#"
(func (export "_start")
  i32.const 0x0F0F0F0F
  i32.popcnt
  call $log
)
"#, "16" },

    test_i32_popcnt_zero => { r#"
(func (export "_start")
  i32.const 0
  i32.popcnt
  call $log
)
"#, "0" },

    test_i32_popcnt_all => { r#"
(func (export "_start")
  i32.const -1
  i32.popcnt
  call $log
)
"#, "32" }
}
