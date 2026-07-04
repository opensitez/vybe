use crate::wat_exec;

wat_exec! {
    test_i64_add_pos_pos => { r#"(func (export "_start") i64.const 10 i64.const 20 i64.add call $log_i64)"#, "30" },
    test_i64_add_pos_neg => { r#"(func (export "_start") i64.const 10 i64.const -5 i64.add call $log_i64)"#, "5" },
    test_i64_add_neg_neg => { r#"(func (export "_start") i64.const -10 i64.const -20 i64.add call $log_i64)"#, "-30" },
    test_i64_add_zero => { r#"(func (export "_start") i64.const 42 i64.const 0 i64.add call $log_i64)"#, "42" },
    test_i64_add_overflow => { r#"(func (export "_start") i64.const 9223372036854775807 i64.const 1 i64.add call $log_i64)"#, "-9223372036854775808" },

    test_i64_sub_pos_pos => { r#"(func (export "_start") i64.const 20 i64.const 10 i64.sub call $log_i64)"#, "10" },
    test_i64_sub_pos_neg => { r#"(func (export "_start") i64.const 10 i64.const -5 i64.sub call $log_i64)"#, "15" },
    test_i64_sub_neg_neg => { r#"(func (export "_start") i64.const -10 i64.const -20 i64.sub call $log_i64)"#, "10" },
    test_i64_sub_zero => { r#"(func (export "_start") i64.const 42 i64.const 0 i64.sub call $log_i64)"#, "42" },
    test_i64_sub_underflow => { r#"(func (export "_start") i64.const -9223372036854775808 i64.const 1 i64.sub call $log_i64)"#, "9223372036854775807" },

    test_i64_mul_pos_pos => { r#"(func (export "_start") i64.const 10 i64.const 20 i64.mul call $log_i64)"#, "200" },
    test_i64_mul_pos_neg => { r#"(func (export "_start") i64.const 10 i64.const -5 i64.mul call $log_i64)"#, "-50" },
    test_i64_mul_neg_neg => { r#"(func (export "_start") i64.const -10 i64.const -20 i64.mul call $log_i64)"#, "200" },
    test_i64_mul_zero => { r#"(func (export "_start") i64.const 42 i64.const 0 i64.mul call $log_i64)"#, "0" },
    test_i64_mul_overflow => { r#"(func (export "_start") i64.const 3037000499 i64.const 3037000499 i64.mul call $log_i64)"#, "9223372030926249001" },

    test_i64_div_s_pos_pos => { r#"(func (export "_start") i64.const 20 i64.const 10 i64.div_s call $log_i64)"#, "2" },
    test_i64_div_s_pos_neg => { r#"(func (export "_start") i64.const 20 i64.const -10 i64.div_s call $log_i64)"#, "-2" },
    test_i64_div_s_neg_pos => { r#"(func (export "_start") i64.const -20 i64.const 10 i64.div_s call $log_i64)"#, "-2" },
    test_i64_div_s_neg_neg => { r#"(func (export "_start") i64.const -20 i64.const -10 i64.div_s call $log_i64)"#, "2" },

    test_i64_div_u_pos_pos => { r#"(func (export "_start") i64.const 20 i64.const 10 i64.div_u call $log_i64)"#, "2" },
    
    test_i64_rem_s_pos_pos => { r#"(func (export "_start") i64.const 20 i64.const 3 i64.rem_s call $log_i64)"#, "2" },
    test_i64_rem_s_pos_neg => { r#"(func (export "_start") i64.const 20 i64.const -3 i64.rem_s call $log_i64)"#, "2" },
    test_i64_rem_s_neg_pos => { r#"(func (export "_start") i64.const -20 i64.const 3 i64.rem_s call $log_i64)"#, "-2" },
    test_i64_rem_s_neg_neg => { r#"(func (export "_start") i64.const -20 i64.const -3 i64.rem_s call $log_i64)"#, "-2" },

    test_i64_rem_u_pos_pos => { r#"(func (export "_start") i64.const 20 i64.const 3 i64.rem_u call $log_i64)"#, "2" }
}
