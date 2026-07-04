use crate::wat_exec;

wat_exec! {
    test_f32_add => { r#"
(func (export "_start")
  f32.const 1.5
  f32.const 2.5
  f32.add
  call $log_f32
)
"#, "4.0" },

    test_f32_sub => { r#"
(func (export "_start")
  f32.const 5.0
  f32.const 2.25
  f32.sub
  call $log_f32
)
"#, "2.75" },

    test_f32_mul => { r#"
(func (export "_start")
  f32.const 3.0
  f32.const 2.5
  f32.mul
  call $log_f32
)
"#, "7.5" },

    test_f32_div => { r#"
(func (export "_start")
  f32.const 10.0
  f32.const 2.5
  f32.div
  call $log_f32
)
"#, "4.0" },

    test_f32_abs_pos => { r#"
(func (export "_start")
  f32.const 3.14
  f32.abs
  call $log_f32
)
"#, "3.14" },

    test_f32_abs_neg => { r#"
(func (export "_start")
  f32.const -3.14
  f32.abs
  call $log_f32
)
"#, "3.14" },

    test_f32_neg => { r#"
(func (export "_start")
  f32.const 3.14
  f32.neg
  call $log_f32
)
"#, "-3.14" },

    test_f32_ceil => { r#"
(func (export "_start")
  f32.const 3.14
  f32.ceil
  call $log_f32
)
"#, "4.0" },
    
    test_f32_ceil_neg => { r#"
(func (export "_start")
  f32.const -3.14
  f32.ceil
  call $log_f32
)
"#, "-3.0" },

    test_f32_floor => { r#"
(func (export "_start")
  f32.const 3.8
  f32.floor
  call $log_f32
)
"#, "3.0" },

    test_f32_floor_neg => { r#"
(func (export "_start")
  f32.const -3.8
  f32.floor
  call $log_f32
)
"#, "-4.0" },

    test_f32_trunc => { r#"
(func (export "_start")
  f32.const 3.8
  f32.trunc
  call $log_f32
)
"#, "3.0" },

    test_f32_trunc_neg => { r#"
(func (export "_start")
  f32.const -3.8
  f32.trunc
  call $log_f32
)
"#, "-3.0" },

    test_f32_nearest_down => { r#"
(func (export "_start")
  f32.const 3.2
  f32.nearest
  call $log_f32
)
"#, "3.0" },

    test_f32_nearest_up => { r#"
(func (export "_start")
  f32.const 3.8
  f32.nearest
  call $log_f32
)
"#, "4.0" },

    test_f32_nearest_half_even => { r#"
(func (export "_start")
  f32.const 2.5
  f32.nearest
  call $log_f32
)
"#, "2.0" },

    test_f32_nearest_half_even_up => { r#"
(func (export "_start")
  f32.const 3.5
  f32.nearest
  call $log_f32
)
"#, "4.0" },

    test_f32_sqrt => { r#"
(func (export "_start")
  f32.const 16.0
  f32.sqrt
  call $log_f32
)
"#, "4.0" },

    test_f32_min => { r#"
(func (export "_start")
  f32.const 3.0
  f32.const 5.0
  f32.min
  call $log_f32
)
"#, "3.0" },

    test_f32_max => { r#"
(func (export "_start")
  f32.const 3.0
  f32.const 5.0
  f32.max
  call $log_f32
)
"#, "5.0" },

    test_f32_copysign => { r#"
(func (export "_start")
  f32.const 3.14
  f32.const -0.0
  f32.copysign
  call $log_f32
)
"#, "-3.14" },

    test_f32_copysign_pos => { r#"
(func (export "_start")
  f32.const -3.14
  f32.const 1.0
  f32.copysign
  call $log_f32
)
"#, "3.14" }
}
