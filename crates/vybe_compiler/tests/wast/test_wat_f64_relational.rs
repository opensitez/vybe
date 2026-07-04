use crate::wat_exec;

wat_exec! {
    test_f64_eq => { r#"
(func (export "_start")
  f64.const 3.14
  f64.const 3.14
  f64.eq
  call $log
)
"#, "1" },

    test_f64_eq_false => { r#"
(func (export "_start")
  f64.const 3.14
  f64.const 2.71
  f64.eq
  call $log
)
"#, "0" },
    
    test_f64_eq_nan => { r#"
(func (export "_start")
  f64.const nan
  f64.const nan
  f64.eq
  call $log
)
"#, "0" }, // NaN != NaN

    test_f64_ne => { r#"
(func (export "_start")
  f64.const 3.14
  f64.const 2.71
  f64.ne
  call $log
)
"#, "1" },

    test_f64_lt => { r#"
(func (export "_start")
  f64.const 2.71
  f64.const 3.14
  f64.lt
  call $log
)
"#, "1" },

    test_f64_le => { r#"
(func (export "_start")
  f64.const 3.14
  f64.const 3.14
  f64.le
  call $log
)
"#, "1" },

    test_f64_gt => { r#"
(func (export "_start")
  f64.const 3.14
  f64.const 2.71
  f64.gt
  call $log
)
"#, "1" },

    test_f64_ge => { r#"
(func (export "_start")
  f64.const 3.14
  f64.const 3.14
  f64.ge
  call $log
)
"#, "1" },

    test_f64_lt_nan => { r#"
(func (export "_start")
  f64.const nan
  f64.const 1.0
  f64.lt
  call $log
)
"#, "0" },

    test_f64_gt_nan => { r#"
(func (export "_start")
  f64.const nan
  f64.const 1.0
  f64.gt
  call $log
)
"#, "0" },

    test_f64_le_nan => { r#"
(func (export "_start")
  f64.const nan
  f64.const 1.0
  f64.le
  call $log
)
"#, "0" },

    test_f64_ge_nan => { r#"
(func (export "_start")
  f64.const nan
  f64.const 1.0
  f64.ge
  call $log
)
"#, "0" }
}
