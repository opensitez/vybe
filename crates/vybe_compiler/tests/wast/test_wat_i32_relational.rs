use crate::wat_exec;

wat_exec! {
    test_i32_eq => { r#"
(func (export "_start")
  i32.const 42
  i32.const 42
  i32.eq
  call $log
)
"#, "1" },

    test_i32_eq_false => { r#"
(func (export "_start")
  i32.const 42
  i32.const 43
  i32.eq
  call $log
)
"#, "0" },

    test_i32_ne => { r#"
(func (export "_start")
  i32.const 42
  i32.const 43
  i32.ne
  call $log
)
"#, "1" },

    test_i32_lt_s => { r#"
(func (export "_start")
  i32.const -1
  i32.const 1
  i32.lt_s
  call $log
)
"#, "1" },

    test_i32_lt_u => { r#"
(func (export "_start")
  i32.const -1 ;; 0xFFFFFFFF
  i32.const 1
  i32.lt_u
  call $log
)
"#, "0" },

    test_i32_le_s => { r#"
(func (export "_start")
  i32.const -1
  i32.const -1
  i32.le_s
  call $log
)
"#, "1" },

    test_i32_le_u => { r#"
(func (export "_start")
  i32.const -1
  i32.const -1
  i32.le_u
  call $log
)
"#, "1" },

    test_i32_gt_s => { r#"
(func (export "_start")
  i32.const 1
  i32.const -1
  i32.gt_s
  call $log
)
"#, "1" },

    test_i32_gt_u => { r#"
(func (export "_start")
  i32.const -1 ;; 0xFFFFFFFF
  i32.const 1
  i32.gt_u
  call $log
)
"#, "1" },

    test_i32_ge_s => { r#"
(func (export "_start")
  i32.const -1
  i32.const -1
  i32.ge_s
  call $log
)
"#, "1" },

    test_i32_ge_u => { r#"
(func (export "_start")
  i32.const -1
  i32.const -1
  i32.ge_u
  call $log
)
"#, "1" },
    
    test_i32_eqz => { r#"
(func (export "_start")
  i32.const 0
  i32.eqz
  call $log
)
"#, "1" },

    test_i32_eqz_false => { r#"
(func (export "_start")
  i32.const 42
  i32.eqz
  call $log
)
"#, "0" }
}
