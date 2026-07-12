//! Extended mod/modulo (real semantics), dim, sign, atan2, hypot, and MERGE on reals.
//! Distinct from `test_integer_mod_division.rs`, `test_trig_hyperbolic.rs`,
//! `test_math_f2008.rs`, and integer MERGE in `test_where_merge_extended.rs`.

fortran_cases! {
    real_mod_pos_pos_nint15 => {
        "program t\nprint *, nint(mod(7.5, 2.0)*10)\nend program t\n",
        ["15"]
    };

    real_mod_neg_pos_nint_neg15 => {
        "program t\nprint *, nint(mod(-7.5, 2.0)*10)\nend program t\n",
        ["-15"]
    };

    real_mod_pos_neg_nint15 => {
        "program t\nprint *, nint(mod(7.5, -2.0)*10)\nend program t\n",
        ["15"]
    };

    real_mod_neg_neg_nint_neg15 => {
        "program t\nprint *, nint(mod(-7.5, -2.0)*10)\nend program t\n",
        ["-15"]
    };

    real_modulo_neg_pos_nint5 => {
        "program t\nprint *, nint(modulo(-7.5, 2.0)*10)\nend program t\n",
        ["5"]
    };

    real_modulo_pos_neg_nint_neg5 => {
        "program t\nprint *, nint(modulo(7.5, -2.0)*10)\nend program t\n",
        ["-5"]
    };

    real_modulo_neg_neg_nint_neg5 => {
        "program t\nprint *, nint(modulo(-7.5, -2.0)*10)\nend program t\n",
        ["-5"]
    };

    real_modulo_pos_pos_nint15 => {
        "program t\nprint *, nint(modulo(7.5, 2.0)*10)\nend program t\n",
        ["15"]
    };

    real_mod_zero_multiple_nint0 => {
        "program t\nprint *, nint(mod(6.0, 3.0)*10)\nend program t\n",
        ["0"]
    };

    real_modulo_zero_multiple_nint0 => {
        "program t\nprint *, nint(modulo(6.0, 3.0)*10)\nend program t\n",
        ["0"]
    };

    real_mod_vs_modulo_neg_dividend => {
        "program t\nprint *, nint(mod(-11.5, 4.0)*10)\nprint *, nint(modulo(-11.5, 4.0)*10)\nend program t\n",
        ["-15", "5"]
    };

    real_mod_vs_modulo_neg_divisor => {
        "program t\nprint *, nint(mod(11.5, -4.0)*10)\nprint *, nint(modulo(11.5, -4.0)*10)\nend program t\n",
        ["15", "-5"]
    };

    do_real_modulo_wraps_angle_degrees => {
        "program t\nreal :: a\ninteger :: i, c\nc = 0\ndo i = 0, 359\na = real(i)\nif (nint(modulo(a, 90.0)) == 0) c = c + 1\nend do\nprint *, c\nend program t\n",
        ["4"]
    };

    do_real_mod_counts_quarter_steps => {
        "program t\nreal :: x\ninteger :: i, c\nc = 0\ndo i = 1, 40\nx = i * 0.25\nif (mod(x, 1.0) == 0.0) c = c + 1\nend do\nprint *, c\nend program t\n",
        ["10"]
    };

    real_mod_reconstructs_dividend => {
        "program t\nreal :: a=29.5, b=6.0, q, r\nq = a / b\nr = mod(a, b)\nprint *, merge(1, 0, q*b + r == a)\nend program t\n",
        ["1"]
    };

    real_modulo_reconstructs_negative => {
        "program t\nreal :: a=-29.5, b=6.0, q, r\nq = a / b\nr = modulo(a, b)\nprint *, merge(1, 0, q*b + r == a)\nend program t\n",
        ["1"]
    };

    dim_int_positive_difference => {
        "program t\nprint *, dim(10, 3)\nend program t\n",
        ["7"]
    };

    dim_int_zero_when_second_larger => {
        "program t\nprint *, dim(3, 10)\nend program t\n",
        ["0"]
    };

    dim_int_equal_operands => {
        "program t\nprint *, dim(5, 5)\nend program t\n",
        ["0"]
    };

    dim_int_negative_first => {
        "program t\nprint *, dim(-2, 5)\nend program t\n",
        ["0"]
    };

    dim_int_both_negative => {
        "program t\nprint *, dim(-8, -3)\nend program t\n",
        ["0"]
    };

    dim_real_scaled_nint73 => {
        "program t\nprint *, nint(dim(10.5, 3.2)*10)\nend program t\n",
        ["73"]
    };

    dim_real_zero_scaled => {
        "program t\nprint *, nint(dim(3.2, 10.5)*10)\nend program t\n",
        ["0"]
    };

    dim_real_equal_scaled => {
        "program t\nprint *, nint(dim(4.0, 4.0)*10)\nend program t\n",
        ["0"]
    };

    dim_in_sum_accumulator => {
        "program t\ninteger :: a(5)=[3,8,1,9,4]\ninteger :: b(5)=[7,2,6,1,5]\nprint *, sum(dim(a,b))\nend program t\n",
        ["12"]
    };

    dim_with_variables => {
        "program t\ninteger :: x=14, y=9\nprint *, dim(x, y)\nend program t\n",
        ["5"]
    };

    sign_int_pos_to_neg => {
        "program t\nprint *, sign(5, -1)\nend program t\n",
        ["-5"]
    };

    sign_int_neg_to_pos => {
        "program t\nprint *, sign(-5, 1)\nend program t\n",
        ["5"]
    };

    sign_int_zero_sign_arg => {
        "program t\nprint *, sign(7, 0)\nend program t\n",
        ["7"]
    };

    sign_int_neg_zero_sign => {
        "program t\nprint *, sign(-7, 0)\nend program t\n",
        ["-7"]
    };

    sign_real_scaled_neg314 => {
        "program t\nprint *, nint(sign(3.14, -1.0)*100)\nend program t\n",
        ["-314"]
    };

    sign_real_scaled_pos271 => {
        "program t\nprint *, nint(sign(2.71, 1.0)*100)\nend program t\n",
        ["271"]
    };

    sign_real_neg_magnitude_pos_sign => {
        "program t\nprint *, nint(sign(-4.5, 2.0)*10)\nend program t\n",
        ["45"]
    };

    sign_real_zero_magnitude => {
        "program t\nprint *, nint(sign(0.0, -9.0)*10)\nend program t\n",
        ["0"]
    };

    sign_preserves_second_arg_sign_negative => {
        "program t\nprint *, sign(100, -3)\nend program t\n",
        ["-100"]
    };

    sign_array_elementwise => {
        "program t\ninteger :: a(3)=[5,-5,0]\ninteger :: s(3)=[-1,1,-1]\nprint *, sign(a(1), s(1))\nprint *, sign(a(2), s(2))\nprint *, sign(a(3), s(3))\nend program t\n",
        ["-5", "5", "0"]
    };

    sign_in_expression_with_dim => {
        "program t\nprint *, sign(dim(8,3), -1)\nend program t\n",
        ["-5"]
    };

    sign_real_with_negative_zero_sign => {
        "program t\nprint *, nint(sign(9.0, -0.0)*10)\nend program t\n",
        ["-90"]
    };

    atan2_east_axis_degrees_zero => {
        "program t\nprint *, nint(atan2(0.0, 1.0)*180/3.14159265)\nend program t\n",
        ["0"]
    };

    atan2_north_axis_degrees_90 => {
        "program t\nprint *, nint(atan2(1.0, 0.0)*180/3.14159265)\nend program t\n",
        ["90"]
    };

    atan2_west_axis_degrees_180 => {
        "program t\nprint *, nint(atan2(0.0, -1.0)*180/3.14159265)\nend program t\n",
        ["180"]
    };

    atan2_south_axis_degrees_neg90 => {
        "program t\nprint *, nint(atan2(-1.0, 0.0)*180/3.14159265)\nend program t\n",
        ["-90"]
    };

    atan2_first_quadrant_45 => {
        "program t\nprint *, nint(atan2(1.0, 1.0)*180/3.14159265)\nend program t\n",
        ["45"]
    };

    atan2_second_quadrant_135 => {
        "program t\nprint *, nint(atan2(1.0, -1.0)*180/3.14159265)\nend program t\n",
        ["135"]
    };

    atan2_third_quadrant_neg135 => {
        "program t\nprint *, nint(atan2(-1.0, -1.0)*180/3.14159265)\nend program t\n",
        ["-135"]
    };

    atan2_fourth_quadrant_neg45 => {
        "program t\nprint *, nint(atan2(-1.0, 1.0)*180/3.14159265)\nend program t\n",
        ["-45"]
    };

    atan2_3_4_triangle_degrees => {
        "program t\nprint *, nint(atan2(3.0, 4.0)*180/3.14159265)\nend program t\n",
        ["37"]
    };

    atan2_neg12_5_obtuse => {
        "program t\nprint *, nint(atan2(5.0, -12.0)*180/3.14159265)\nend program t\n",
        ["157"]
    };

    hypot_3_4_is_5 => {
        "program t\nprint *, nint(hypot(3.0, 4.0))\nend program t\n",
        ["5"]
    };

    hypot_5_12_is_13 => {
        "program t\nprint *, nint(hypot(5.0, 12.0))\nend program t\n",
        ["13"]
    };

    hypot_8_15_is_17 => {
        "program t\nprint *, nint(hypot(8.0, 15.0))\nend program t\n",
        ["17"]
    };

    hypot_x_zero_y_only => {
        "program t\nprint *, nint(hypot(0.0, 9.0))\nend program t\n",
        ["9"]
    };

    hypot_y_zero_x_only => {
        "program t\nprint *, nint(hypot(7.0, 0.0))\nend program t\n",
        ["7"]
    };

    hypot_both_zero => {
        "program t\nprint *, nint(hypot(0.0, 0.0))\nend program t\n",
        ["0"]
    };

    hypot_negative_legs_same_as_positive => {
        "program t\nprint *, nint(hypot(-3.0, -4.0))\nend program t\n",
        ["5"]
    };

    hypot_scaled_triangle => {
        "program t\nprint *, nint(hypot(30.0, 40.0))\nend program t\n",
        ["50"]
    };

    merge_real_scalar_true_branch => {
        "program t\nreal :: x\nx = merge(3.5, 7.5, .true.)\nprint *, nint(x*10)\nend program t\n",
        ["35"]
    };

    merge_real_scalar_false_branch => {
        "program t\nreal :: x\nx = merge(3.5, 7.5, .false.)\nprint *, nint(x*10)\nend program t\n",
        ["75"]
    };

    merge_real_array_by_mask => {
        "program t\nreal :: a(3)=[1.5,2.5,3.5]\nreal :: b(3)=[9.0,8.0,7.0]\nreal :: c(3)\nlogical :: m(3)=[.true.,.false.,.true.]\nc = merge(a, b, m)\nprint *, nint(c(1)*10)\nprint *, nint(c(2)*10)\nprint *, nint(c(3)*10)\nend program t\n",
        ["15", "80", "35"]
    };

    merge_real_with_negative_values => {
        "program t\nreal :: x\nx = merge(-2.5, 4.0, .false.)\nprint *, nint(x*10)\nend program t\n",
        ["40"]
    };

    merge_real_nested_in_expression => {
        "program t\nreal :: x\nx = merge(merge(1.0,2.0,.true.), merge(3.0,4.0,.false.), .true.)\nprint *, nint(x*10)\nend program t\n",
        ["10"]
    };

    merge_real_do_loop_clamp_negatives => {
        "program t\nreal :: v(4)=[-1.5,2.0,-3.0,4.0]\nreal :: w(4)\ninteger :: i\ndo i=1,4\nw(i) = merge(v(i), 0.0, v(i)>0.0)\nend do\nprint *, nint(w(1)*10)\nprint *, nint(w(2)*10)\nprint *, nint(w(4)*10)\nend program t\n",
        ["0", "20", "40"]
    };

    merge_real_abs_via_sign => {
        "program t\nreal :: x\nx = merge(-6.0, 6.0, .false.)\nprint *, nint(x)\nend program t\n",
        ["6"]
    };

    merge_real_with_logical_from_compare => {
        "program t\nreal :: a=2.0, b=5.0\nprint *, nint(merge(a, b, a<b)*10)\nend program t\n",
        ["20"]
    };

    merge_real_2d_slice => {
        "program t\nreal :: a(2,2)=reshape([1.0,2.0,3.0,4.0],[2,2])\nreal :: b(2,2)=reshape([10.0,20.0,30.0,40.0],[2,2])\nreal :: c(2,2)\nc = merge(a, b, a<3.0)\nprint *, nint(c(1,1)*10)\nprint *, nint(c(2,2)*10)\nend program t\n",
        ["10", "40"]
    };

    merge_real_sum_selected => {
        "program t\nreal :: a(3)=[0.5,1.5,2.5]\nreal :: b(3)=[5.0,4.0,3.0]\nlogical :: m(3)=[.true.,.false.,.true.]\nprint *, nint(sum(merge(a,b,m))*10)\nend program t\n",
        ["90"]
    };

}
