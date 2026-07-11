use crate::wat_exec;

wat_exec! {
    test_f32_eq => { r#"
(func (export "_start")
  f32.const 3.14
  f32.const 3.14
  f32.eq
  call $log
)
"#, "1" },

    test_f32_eq_false => { r#"
(func (export "_start")
  f32.const 3.14
  f32.const 2.71
  f32.eq
  call $log
)
"#, "0" },

    test_f32_eq_nan => { r#"
(func (export "_start")
  f32.const nan
  f32.const nan
  f32.eq
  call $log
)
"#, "0" }, // NaN != NaN

    test_f32_ne => { r#"
(func (export "_start")
  f32.const 3.14
  f32.const 2.71
  f32.ne
  call $log
)
"#, "1" },

    test_f32_lt => { r#"
(func (export "_start")
  f32.const 2.71
  f32.const 3.14
  f32.lt
  call $log
)
"#, "1" },

    test_f32_le => { r#"
(func (export "_start")
  f32.const 3.14
  f32.const 3.14
  f32.le
  call $log
)
"#, "1" },

    test_f32_gt => { r#"
(func (export "_start")
  f32.const 3.14
  f32.const 2.71
  f32.gt
  call $log
)
"#, "1" },

    test_f32_ge => { r#"
(func (export "_start")
  f32.const 3.14
  f32.const 3.14
  f32.ge
  call $log
)
"#, "1" },

    test_f32_lt_nan => { r#"
(func (export "_start")
  f32.const nan
  f32.const 1.0
  f32.lt
  call $log
)
"#, "0" },

    test_f32_gt_nan => { r#"
(func (export "_start")
  f32.const nan
  f32.const 1.0
  f32.gt
  call $log
)
"#, "0" },

    test_f32_le_nan => { r#"
(func (export "_start")
  f32.const nan
  f32.const 1.0
  f32.le
  call $log
)
"#, "0" },

    test_f32_ge_nan => { r#"
(func (export "_start")
  f32.const nan
  f32.const 1.0
  f32.ge
  call $log
)
"#, "0" }
}
