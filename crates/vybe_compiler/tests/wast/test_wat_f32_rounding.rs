use crate::wat_exec;

wat_exec! {
    test_f32_ceil_pos => { r#"(func (export "_start") f32.const 1.2 f32.ceil call $log_f32)"#, "2.0" },
    test_f32_ceil_exact => { r#"(func (export "_start") f32.const 2.0 f32.ceil call $log_f32)"#, "2.0" },
    test_f32_ceil_neg => { r#"(func (export "_start") f32.const -1.2 f32.ceil call $log_f32)"#, "-1.0" },
    test_f32_ceil_nan => { r#"(func (export "_start") f32.const nan f32.ceil call $log_f32)"#, "nan" },

    test_f32_floor_pos => { r#"(func (export "_start") f32.const 1.8 f32.floor call $log_f32)"#, "1.0" },
    test_f32_floor_exact => { r#"(func (export "_start") f32.const 2.0 f32.floor call $log_f32)"#, "2.0" },
    test_f32_floor_neg => { r#"(func (export "_start") f32.const -1.8 f32.floor call $log_f32)"#, "-2.0" },
    test_f32_floor_inf => { r#"(func (export "_start") f32.const inf f32.floor call $log_f32)"#, "inf" },

    test_f32_trunc_pos => { r#"(func (export "_start") f32.const 1.8 f32.trunc call $log_f32)"#, "1.0" },
    test_f32_trunc_exact => { r#"(func (export "_start") f32.const 2.0 f32.trunc call $log_f32)"#, "2.0" },
    test_f32_trunc_neg => { r#"(func (export "_start") f32.const -1.8 f32.trunc call $log_f32)"#, "-1.0" },

    test_f32_nearest_pos_down => { r#"(func (export "_start") f32.const 1.2 f32.nearest call $log_f32)"#, "1.0" },
    test_f32_nearest_pos_up => { r#"(func (export "_start") f32.const 1.8 f32.nearest call $log_f32)"#, "2.0" },
    test_f32_nearest_pos_half_even => { r#"(func (export "_start") f32.const 1.5 f32.nearest call $log_f32)"#, "2.0" },
    test_f32_nearest_pos_half_even_2 => { r#"(func (export "_start") f32.const 2.5 f32.nearest call $log_f32)"#, "2.0" },
    
    test_f32_nearest_neg_down => { r#"(func (export "_start") f32.const -1.8 f32.nearest call $log_f32)"#, "-2.0" },
    test_f32_nearest_neg_up => { r#"(func (export "_start") f32.const -1.2 f32.nearest call $log_f32)"#, "-1.0" },
    test_f32_nearest_neg_half_even => { r#"(func (export "_start") f32.const -1.5 f32.nearest call $log_f32)"#, "-2.0" },
    test_f32_nearest_neg_half_even_2 => { r#"(func (export "_start") f32.const -2.5 f32.nearest call $log_f32)"#, "-2.0" }
}
