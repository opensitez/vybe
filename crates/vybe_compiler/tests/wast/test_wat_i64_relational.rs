use crate::wat_exec;

wat_exec! {
    test_i64_eq => { r#"
(func (export "_start")
  i64.const 42
  i64.const 42
  i64.eq
  call $log
)
"#, "1" },

    test_i64_eq_false => { r#"
(func (export "_start")
  i64.const 42
  i64.const 43
  i64.eq
  call $log
)
"#, "0" },

    test_i64_ne => { r#"
(func (export "_start")
  i64.const 42
  i64.const 43
  i64.ne
  call $log
)
"#, "1" },

    test_i64_lt_s => { r#"
(func (export "_start")
  i64.const -1
  i64.const 1
  i64.lt_s
  call $log
)
"#, "1" },

    test_i64_lt_u => { r#"
(func (export "_start")
  i64.const -1
  i64.const 1
  i64.lt_u
  call $log
)
"#, "0" },

    test_i64_le_s => { r#"
(func (export "_start")
  i64.const -1
  i64.const -1
  i64.le_s
  call $log
)
"#, "1" },

    test_i64_le_u => { r#"
(func (export "_start")
  i64.const -1
  i64.const -1
  i64.le_u
  call $log
)
"#, "1" },

    test_i64_gt_s => { r#"
(func (export "_start")
  i64.const 1
  i64.const -1
  i64.gt_s
  call $log
)
"#, "1" },

    test_i64_gt_u => { r#"
(func (export "_start")
  i64.const -1
  i64.const 1
  i64.gt_u
  call $log
)
"#, "1" },

    test_i64_ge_s => { r#"
(func (export "_start")
  i64.const -1
  i64.const -1
  i64.ge_s
  call $log
)
"#, "1" },

    test_i64_ge_u => { r#"
(func (export "_start")
  i64.const -1
  i64.const -1
  i64.ge_u
  call $log
)
"#, "1" },
    
    test_i64_eqz => { r#"
(func (export "_start")
  i64.const 0
  i64.eqz
  call $log
)
"#, "1" },

    test_i64_eqz_false => { r#"
(func (export "_start")
  i64.const 42
  i64.eqz
  call $log
)
"#, "0" }
}
