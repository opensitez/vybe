use crate::wat_exec;

wat_exec! {
    test_f64_add => { r#"
(func (export "_start")
  f64.const 1.5
  f64.const 2.5
  f64.add
  call $log_f64
)
"#, "4.0" },

    test_f64_sub => { r#"
(func (export "_start")
  f64.const 5.0
  f64.const 2.25
  f64.sub
  call $log_f64
)
"#, "2.75" },

    test_f64_mul => { r#"
(func (export "_start")
  f64.const 3.0
  f64.const 2.5
  f64.mul
  call $log_f64
)
"#, "7.5" },

    test_f64_div => { r#"
(func (export "_start")
  f64.const 10.0
  f64.const 2.5
  f64.div
  call $log_f64
)
"#, "4.0" },

    test_f64_abs_pos => { r#"
(func (export "_start")
  f64.const 3.14
  f64.abs
  call $log_f64
)
"#, "3.14" },

    test_f64_abs_neg => { r#"
(func (export "_start")
  f64.const -3.14
  f64.abs
  call $log_f64
)
"#, "3.14" },

    test_f64_neg => { r#"
(func (export "_start")
  f64.const 3.14
  f64.neg
  call $log_f64
)
"#, "-3.14" },

    test_f64_ceil => { r#"
(func (export "_start")
  f64.const 3.14
  f64.ceil
  call $log_f64
)
"#, "4.0" },

    test_f64_ceil_neg => { r#"
(func (export "_start")
  f64.const -3.14
  f64.ceil
  call $log_f64
)
"#, "-3.0" },

    test_f64_floor => { r#"
(func (export "_start")
  f64.const 3.8
  f64.floor
  call $log_f64
)
"#, "3.0" },

    test_f64_floor_neg => { r#"
(func (export "_start")
  f64.const -3.8
  f64.floor
  call $log_f64
)
"#, "-4.0" },

    test_f64_trunc => { r#"
(func (export "_start")
  f64.const 3.8
  f64.trunc
  call $log_f64
)
"#, "3.0" },

    test_f64_trunc_neg => { r#"
(func (export "_start")
  f64.const -3.8
  f64.trunc
  call $log_f64
)
"#, "-3.0" },

    test_f64_nearest_down => { r#"
(func (export "_start")
  f64.const 3.2
  f64.nearest
  call $log_f64
)
"#, "3.0" },

    test_f64_nearest_up => { r#"
(func (export "_start")
  f64.const 3.8
  f64.nearest
  call $log_f64
)
"#, "4.0" },

    test_f64_nearest_half_even => { r#"
(func (export "_start")
  f64.const 2.5
  f64.nearest
  call $log_f64
)
"#, "2.0" },

    test_f64_nearest_half_even_up => { r#"
(func (export "_start")
  f64.const 3.5
  f64.nearest
  call $log_f64
)
"#, "4.0" },

    test_f64_sqrt => { r#"
(func (export "_start")
  f64.const 16.0
  f64.sqrt
  call $log_f64
)
"#, "4.0" },

    test_f64_min => { r#"
(func (export "_start")
  f64.const 3.0
  f64.const 5.0
  f64.min
  call $log_f64
)
"#, "3.0" },

    test_f64_max => { r#"
(func (export "_start")
  f64.const 3.0
  f64.const 5.0
  f64.max
  call $log_f64
)
"#, "5.0" },

    test_f64_copysign => { r#"
(func (export "_start")
  f64.const 3.14
  f64.const -0.0
  f64.copysign
  call $log_f64
)
"#, "-3.14" },

    test_f64_copysign_pos => { r#"
(func (export "_start")
  f64.const -3.14
  f64.const 1.0
  f64.copysign
  call $log_f64
)
"#, "3.14" }
}
