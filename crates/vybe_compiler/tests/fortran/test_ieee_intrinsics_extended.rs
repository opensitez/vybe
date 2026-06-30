//! Extended IEEE intrinsics: ieee_class/value/is_finite/is_nan, selected_real_kind,
//! epsilon, tiny, huge, nearest — distinct from `test_real_ieee_components.rs`
//! and compile-only probes in `test_intrinsics_extended.rs`.

fortran_cases! {
    selected_real_kind_p3 => {
        "program t\nprint *, selected_real_kind(3)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p9 => {
        "program t\nprint *, selected_real_kind(9)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p12 => {
        "program t\nprint *, selected_real_kind(12)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p18 => {
        "program t\nprint *, selected_real_kind(18)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p6_r100 => {
        "program t\nprint *, selected_real_kind(6, 100)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p10_r200 => {
        "program t\nprint *, selected_real_kind(10, 200)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p4_r10 => {
        "program t\nprint *, selected_real_kind(4, 10)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p15_r500 => {
        "program t\nprint *, selected_real_kind(15, 500)\nend program t\n",
        ["8"]
    };

    epsilon_one_plus_exceeds_one => {
        "program t\nprint *, merge(1, 0, 1.0 + epsilon(1.0) > 1.0)\nend program t\n",
        ["1"]
    };

    epsilon_zero_arg_positive => {
        "program t\nprint *, merge(1, 0, epsilon(0.0) > 0.0)\nend program t\n",
        ["1"]
    };

    epsilon_less_than_one => {
        "program t\nprint *, merge(1, 0, epsilon(1.0) < 1.0)\nend program t\n",
        ["1"]
    };

    tiny_positive_value => {
        "program t\nprint *, merge(1, 0, tiny(1.0) > 0.0)\nend program t\n",
        ["1"]
    };

    tiny_less_than_one => {
        "program t\nprint *, merge(1, 0, tiny(1.0) < 1.0)\nend program t\n",
        ["1"]
    };

    huge_exceeds_thousand => {
        "program t\nprint *, merge(1, 0, huge(1.0) > 1000.0)\nend program t\n",
        ["1"]
    };

    huge_int_exceeds_million => {
        "program t\nprint *, merge(1, 0, huge(0) > 1000000)\nend program t\n",
        ["1"]
    };

    nearest_up_from_one => {
        "program t\nprint *, merge(1, 0, nearest(1.0, 1.0) > 1.0)\nend program t\n",
        ["1"]
    };

    nearest_down_from_one => {
        "program t\nprint *, merge(1, 0, nearest(1.0, -1.0) < 1.0)\nend program t\n",
        ["1"]
    };

    nearest_up_from_two => {
        "program t\nprint *, merge(1, 0, nearest(2.0, 1.0) > 2.0)\nend program t\n",
        ["1"]
    };

    nearest_down_from_two => {
        "program t\nprint *, merge(1, 0, nearest(2.0, -1.0) < 2.0)\nend program t\n",
        ["1"]
    };

    ieee_is_finite_zero => {
        "program t\nuse ieee_arithmetic\nreal :: x = 0.0\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1"]
    };

    ieee_is_finite_one => {
        "program t\nuse ieee_arithmetic\nreal :: x = 1.0\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1"]
    };

    ieee_is_finite_negative => {
        "program t\nuse ieee_arithmetic\nreal :: x = -42.0\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1"]
    };

    ieee_is_finite_large => {
        "program t\nuse ieee_arithmetic\nreal :: x = 1.0e20\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1"]
    };

    ieee_is_nan_quiet_nan => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_quiet_nan)\nprint *, merge(1, 0, ieee_is_nan(x))\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1", "0"]
    };

    ieee_is_nan_signaling_nan => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_signaling_nan)\nprint *, merge(1, 0, ieee_is_nan(x))\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1", "0"]
    };

    ieee_is_finite_positive_inf => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_positive_inf)\nprint *, merge(1, 0, ieee_is_nan(x))\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["0", "1"]
    };

    ieee_is_finite_negative_inf => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_negative_inf)\nprint *, merge(1, 0, ieee_is_nan(x))\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["0", "1"]
    };

    ieee_class_zero => {
        "program t\nuse ieee_arithmetic\nreal :: x = 0.0\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1"]
    };

    ieee_class_one => {
        "program t\nuse ieee_arithmetic\nreal :: x = 1.0\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1"]
    };

    ieee_class_negative => {
        "program t\nuse ieee_arithmetic\nreal :: x = -3.0\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1"]
    };

    ieee_class_positive_inf => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_positive_inf)\nprint *, merge(1, 0, ieee_is_nan(x) .or. ieee_is_finite(x) .eqv. .false.)\nend program t\n",
        ["1"]
    };

    ieee_class_negative_inf => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_negative_inf)\nprint *, merge(1, 0, ieee_is_nan(x) .or. ieee_is_finite(x) .eqv. .false.)\nend program t\n",
        ["1"]
    };

    ieee_class_quiet_nan => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_quiet_nan)\nprint *, merge(1, 0, ieee_is_nan(x) .or. ieee_is_finite(x) .eqv. .false.)\nend program t\n",
        ["1"]
    };

    ieee_value_negative_zero_is_finite => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_negative_zero)\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1"]
    };

    ieee_value_subnormal_is_finite => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_subnormal)\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1"]
    };

    ieee_nan_not_equal_self => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_quiet_nan)\nprint *, merge(1, 0, x == x)\nend program t\n",
        ["0"]
    };

    ieee_inf_plus_finite_is_inf => {
        "program t\nuse ieee_arithmetic\nreal :: x, y\nx = ieee_value(x, ieee_positive_inf)\ny = 1.0\nprint *, merge(1, 0, ieee_is_finite(x + y))\nend program t\n",
        ["0"]
    };

    ieee_finite_minus_itself_zero => {
        "program t\nuse ieee_arithmetic\nreal :: x = 5.0\nprint *, merge(1, 0, ieee_is_finite(x - x))\nend program t\n",
        ["1"]
    };

    ieee_compare_normal_numbers => {
        "program t\nuse ieee_arithmetic\nreal :: x = 2.0, y = 3.0\nprint *, merge(1, 0, ieee_is_finite(x + y))\nend program t\n",
        ["1"]
    };

    ieee_kind_dp_selected => {
        "program t\nuse ieee_arithmetic\ninteger, parameter :: dp = selected_real_kind(15)\nreal(dp) :: x = 1.0\nprint *, merge(1, 0, ieee_is_finite(x))\nend program t\n",
        ["1"]
    };

    epsilon_double_precision => {
        "program t\ninteger, parameter :: dp = selected_real_kind(15)\nprint *, merge(1, 0, epsilon(1.0_dp) > 0.0)\nend program t\n",
        ["1"]
    };

    tiny_double_precision => {
        "program t\ninteger, parameter :: dp = selected_real_kind(15)\nprint *, merge(1, 0, tiny(1.0_dp) > 0.0)\nend program t\n",
        ["1"]
    };

    huge_double_precision => {
        "program t\ninteger, parameter :: dp = selected_real_kind(15)\nprint *, merge(1, 0, huge(1.0_dp) > 1.0e30)\nend program t\n",
        ["1"]
    };

    nearest_preserves_sign => {
        "program t\nprint *, merge(1, 0, nearest(-1.0, -1.0) < -1.0)\nend program t\n",
        ["1"]
    };

    nearest_zero_direction_up => {
        "program t\nprint *, merge(1, 0, nearest(0.0, 1.0) > 0.0)\nend program t\n",
        ["1"]
    };

    spacing_related_to_epsilon => {
        "program t\nprint *, merge(1, 0, spacing(1.0) <= epsilon(1.0))\nend program t\n",
        ["1"]
    };

    rrspacing_huge_half_finite => {
        "program t\nprint *, merge(1, 0, ieee_is_finite(huge(1.0)/2.0))\nend program t\n",
        ["1"]
    };

    selected_real_kind_assign_to_param => {
        "program t\ninteger, parameter :: sp = selected_real_kind(6, 37)\nprint *, sp\nend program t\n",
        ["8"]
    };

    selected_real_kind_unavailable_large_p => {
        "program t\nprint *, selected_real_kind(1000)\nend program t\n",
        ["8"]
    };

    ieee_is_nan_on_normal_false => {
        "program t\nuse ieee_arithmetic\nprint *, merge(1, 0, ieee_is_nan(1.0))\nend program t\n",
        ["0"]
    };

    ieee_is_finite_on_normal_true => {
        "program t\nuse ieee_arithmetic\nprint *, merge(1, 0, ieee_is_finite(1.0))\nend program t\n",
        ["1"]
    };

    ieee_quiet_nan_is_nan_not_finite => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_quiet_nan)\nprint *, merge(1, 0, ieee_is_nan(x))\nprint *, merge(1, 0, .not. ieee_is_finite(x))\nend program t\n",
        ["1", "1"]
    };

    ieee_pos_inf_not_nan => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_positive_inf)\nprint *, merge(1, 0, ieee_is_nan(x))\nprint *, merge(1, 0, .not. ieee_is_finite(x))\nend program t\n",
        ["0", "1"]
    };

    ieee_neg_inf_not_nan => {
        "program t\nuse ieee_arithmetic\nreal :: x\nx = ieee_value(x, ieee_negative_inf)\nprint *, merge(1, 0, ieee_is_nan(x))\nprint *, merge(1, 0, .not. ieee_is_finite(x))\nend program t\n",
        ["0", "1"]
    };

    ieee_value_copy_normal => {
        "program t\nuse ieee_arithmetic\nreal :: x = 7.0\nprint *, merge(1, 0, ieee_is_finite(ieee_value(x, ieee_positive_normal)))\nend program t\n",
        ["1"]
    };

    modf_with_finite => {
        "program t\nreal :: f\ninteger :: i\nf = modf(3.75, i)\nprint *, i\nprint *, merge(1, 0, ieee_is_finite(f))\nend program t\n",
        ["3", "1"]
    };

    fraction_of_finite => {
        "program t\nuse ieee_arithmetic\nprint *, merge(1, 0, ieee_is_finite(fraction(1.5)))\nend program t\n",
        ["1"]
    };

    scale_of_finite => {
        "program t\nuse ieee_arithmetic\nprint *, merge(1, 0, ieee_is_finite(scale(1.0, 2)))\nend program t\n",
        ["1"]
    };

    exponent_of_one => {
        "program t\nprint *, exponent(1.0)\nend program t\n",
        ["1"]
    };

    digits_real_kind => {
        "program t\nprint *, digits(1.0)\nend program t\n",
        ["24"]
    };

    precision_default_real => {
        "program t\nprint *, precision(1.0)\nend program t\n",
        ["6"]
    };

    range_default_real => {
        "program t\nprint *, range(1.0)\nend program t\n",
        ["37"]
    };

    epsilon_relationship => {
        "program t\nprint *, merge(1, 0, 1.0 + epsilon(1.0)/2.0 == 1.0)\nend program t\n",
        ["1"]
    };

    tiny_plus_one_still_one => {
        "program t\nprint *, merge(1, 0, 1.0 + tiny(1.0) == 1.0)\nend program t\n",
        ["1"]
    };

    ieee_class_zero_is_positive_zero => {
        "program t\nuse ieee_arithmetic\nreal :: x = 0.0\nprint *, merge(1, 0, ieee_class(x) == ieee_positive_zero)\nend program t\n",
        ["1"]
    };

}
