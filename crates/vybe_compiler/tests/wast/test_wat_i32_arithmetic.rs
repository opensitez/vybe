use crate::wat_exec;

wat_exec! {
    test_i32_add_pos_pos => { r#"(func (export "_start") i32.const 10 i32.const 20 i32.add call $log)"#, "30" },
    test_i32_add_pos_neg => { r#"(func (export "_start") i32.const 10 i32.const -5 i32.add call $log)"#, "5" },
    test_i32_add_neg_neg => { r#"(func (export "_start") i32.const -10 i32.const -20 i32.add call $log)"#, "-30" },
    test_i32_add_zero => { r#"(func (export "_start") i32.const 42 i32.const 0 i32.add call $log)"#, "42" },
    test_i32_add_overflow => { r#"(func (export "_start") i32.const 2147483647 i32.const 1 i32.add call $log)"#, "-2147483648" },

    test_i32_sub_pos_pos => { r#"(func (export "_start") i32.const 20 i32.const 10 i32.sub call $log)"#, "10" },
    test_i32_sub_pos_neg => { r#"(func (export "_start") i32.const 10 i32.const -5 i32.sub call $log)"#, "15" },
    test_i32_sub_neg_neg => { r#"(func (export "_start") i32.const -10 i32.const -20 i32.sub call $log)"#, "10" },
    test_i32_sub_zero => { r#"(func (export "_start") i32.const 42 i32.const 0 i32.sub call $log)"#, "42" },
    test_i32_sub_underflow => { r#"(func (export "_start") i32.const -2147483648 i32.const 1 i32.sub call $log)"#, "2147483647" },

    test_i32_mul_pos_pos => { r#"(func (export "_start") i32.const 10 i32.const 20 i32.mul call $log)"#, "200" },
    test_i32_mul_pos_neg => { r#"(func (export "_start") i32.const 10 i32.const -5 i32.mul call $log)"#, "-50" },
    test_i32_mul_neg_neg => { r#"(func (export "_start") i32.const -10 i32.const -20 i32.mul call $log)"#, "200" },
    test_i32_mul_zero => { r#"(func (export "_start") i32.const 42 i32.const 0 i32.mul call $log)"#, "0" },
    test_i32_mul_overflow => { r#"(func (export "_start") i32.const 1000000 i32.const 1000000 i32.mul call $log)"#, "-727379968" },

    test_i32_div_s_pos_pos => { r#"(func (export "_start") i32.const 20 i32.const 10 i32.div_s call $log)"#, "2" },
    test_i32_div_s_pos_neg => { r#"(func (export "_start") i32.const 20 i32.const -10 i32.div_s call $log)"#, "-2" },
    test_i32_div_s_neg_pos => { r#"(func (export "_start") i32.const -20 i32.const 10 i32.div_s call $log)"#, "-2" },
    test_i32_div_s_neg_neg => { r#"(func (export "_start") i32.const -20 i32.const -10 i32.div_s call $log)"#, "2" },

    test_i32_div_u_pos_pos => { r#"(func (export "_start") i32.const 20 i32.const 10 i32.div_u call $log)"#, "2" },

    test_i32_rem_s_pos_pos => { r#"(func (export "_start") i32.const 20 i32.const 3 i32.rem_s call $log)"#, "2" },
    test_i32_rem_s_pos_neg => { r#"(func (export "_start") i32.const 20 i32.const -3 i32.rem_s call $log)"#, "2" },
    test_i32_rem_s_neg_pos => { r#"(func (export "_start") i32.const -20 i32.const 3 i32.rem_s call $log)"#, "-2" },
    test_i32_rem_s_neg_neg => { r#"(func (export "_start") i32.const -20 i32.const -3 i32.rem_s call $log)"#, "-2" },

    test_i32_rem_u_pos_pos => { r#"(func (export "_start") i32.const 20 i32.const 3 i32.rem_u call $log)"#, "2" }
}
