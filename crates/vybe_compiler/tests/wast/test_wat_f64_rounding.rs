use crate::wat_exec;

wat_exec! {
    test_f64_ceil_pos => { r#"(func (export "_start") f64.const 1.2 f64.ceil call $log_f64)"#, "2.0" },
    test_f64_ceil_exact => { r#"(func (export "_start") f64.const 2.0 f64.ceil call $log_f64)"#, "2.0" },
    test_f64_ceil_neg => { r#"(func (export "_start") f64.const -1.2 f64.ceil call $log_f64)"#, "-1.0" },
    test_f64_ceil_nan => { r#"(func (export "_start") f64.const nan f64.ceil call $log_f64)"#, "nan" },

    test_f64_floor_pos => { r#"(func (export "_start") f64.const 1.8 f64.floor call $log_f64)"#, "1.0" },
    test_f64_floor_exact => { r#"(func (export "_start") f64.const 2.0 f64.floor call $log_f64)"#, "2.0" },
    test_f64_floor_neg => { r#"(func (export "_start") f64.const -1.8 f64.floor call $log_f64)"#, "-2.0" },
    test_f64_floor_inf => { r#"(func (export "_start") f64.const inf f64.floor call $log_f64)"#, "inf" },

    test_f64_trunc_pos => { r#"(func (export "_start") f64.const 1.8 f64.trunc call $log_f64)"#, "1.0" },
    test_f64_trunc_exact => { r#"(func (export "_start") f64.const 2.0 f64.trunc call $log_f64)"#, "2.0" },
    test_f64_trunc_neg => { r#"(func (export "_start") f64.const -1.8 f64.trunc call $log_f64)"#, "-1.0" },

    test_f64_nearest_pos_down => { r#"(func (export "_start") f64.const 1.2 f64.nearest call $log_f64)"#, "1.0" },
    test_f64_nearest_pos_up => { r#"(func (export "_start") f64.const 1.8 f64.nearest call $log_f64)"#, "2.0" },
    test_f64_nearest_pos_half_even => { r#"(func (export "_start") f64.const 1.5 f64.nearest call $log_f64)"#, "2.0" },
    test_f64_nearest_pos_half_even_2 => { r#"(func (export "_start") f64.const 2.5 f64.nearest call $log_f64)"#, "2.0" },

    test_f64_nearest_neg_down => { r#"(func (export "_start") f64.const -1.8 f64.nearest call $log_f64)"#, "-2.0" },
    test_f64_nearest_neg_up => { r#"(func (export "_start") f64.const -1.2 f64.nearest call $log_f64)"#, "-1.0" },
    test_f64_nearest_neg_half_even => { r#"(func (export "_start") f64.const -1.5 f64.nearest call $log_f64)"#, "-2.0" },
    test_f64_nearest_neg_half_even_2 => { r#"(func (export "_start") f64.const -2.5 f64.nearest call $log_f64)"#, "-2.0" }
}
