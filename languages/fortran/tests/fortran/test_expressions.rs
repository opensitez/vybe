fortran_cases! {
    expr_add_01 => {
        "program p\ninteger :: a=1,b=2,c\nc = a + b\nprint *, c\nend program p\n",
        ["3"]
    };

    expr_sub_02 => {
        "program p\ninteger :: a=5,b=2,c\nc = a - b\nprint *, c\nend program p\n",
        ["3"]
    };

    expr_mul_03 => {
        "program p\ninteger :: a=3,b=4,c\nc = a * b\nprint *, c\nend program p\n",
        ["12"]
    };

    expr_div_04 => {
        "program p\nreal :: a=8.0,b=2.0,c\nc = a / b\nprint *, c\nend program p\n",
        ["4"]
    };

    expr_pow_05 => {
        "program p\ninteger :: a=2,b\nb = a ** 3\nprint *, b\nend program p\n",
        ["8"]
    };

    expr_unary_06 => {
        "program p\ninteger :: a\na = -5\nprint *, a\nend program p\n",
        ["-5"]
    };

    expr_paren_07 => {
        "program p\ninteger :: x\nx = (2 + 3) * 4\nprint *, x\nend program p\n",
        ["20"]
    };

    expr_prec_08 => {
        "program p\ninteger :: x\nx = 2 + 3 * 4\nprint *, x\nend program p\n",
        ["14"]
    };

    expr_logical_and_09 => {
        "program p\nlogical :: x\nx = .true. .and. .false.\nprint *, x\nend program p\n",
        ["false"]
    };

    expr_logical_or_10 => {
        "program p\nlogical :: x\nx = .true. .or. .false.\nprint *, x\nend program p\n",
        ["true"]
    };

    expr_logical_not_11 => {
        "program p\nlogical :: x\nx = .not. .false.\nprint *, x\nend program p\n",
        ["true"]
    };

    expr_eq_12 => {
        "program p\nlogical :: x\nx = 1 == 1\nprint *, x\nend program p\n",
        ["true"]
    };

    expr_ne_13 => {
        "program p\nlogical :: x\nx = 1 /= 2\nprint *, x\nend program p\n",
        ["true"]
    };

    expr_lt_14 => {
        "program p\nlogical :: x\nx = 1 < 2\nprint *, x\nend program p\n",
        ["true"]
    };

    expr_le_15 => {
        "program p\nlogical :: x\nx = 1 <= 2\nprint *, x\nend program p\n",
        ["true"]
    };

    expr_gt_16 => {
        "program p\nlogical :: x\nx = 2 > 1\nprint *, x\nend program p\n",
        ["true"]
    };

    expr_ge_17 => {
        "program p\nlogical :: x\nx = 2 >= 1\nprint *, x\nend program p\n",
        ["true"]
    };

    expr_concat_18 => {
        "program p\ncharacter(len=2) :: s\ns = 'a'//'b'\nprint *, s\nend program p\n",
        ["ab"]
    };

    expr_char_rel_19 => {
        "program p\nlogical :: x\nx = 'a' < 'b'\nprint *, x\nend program p\n",
        ["true"]
    };

    expr_complex_add_20 => {
        "program p\ncomplex :: a=(1.0,2.0), b=(3.0,4.0), c\nc = a + b\nprint *, real(c)\nprint *, aimag(c)\nend program p\n",
        ["4", "6"]
    };

    expr_array_constructor_21 => {
        "program p\ninteger :: a(3)\na = [1,2,3]\nprint *, a(1) + a(2) + a(3)\nend program p\n",
        ["6"]
    };

    expr_section_22 => {
        "program p\ninteger :: a(4)\na = [1,2,3,4]\nprint *, a(2) + a(3)\nend program p\n",
        ["5"]
    };

    expr_index_23 => {
        "program p\ninteger :: a(3)\na = [1,2,3]\nprint *, a(2)\nend program p\n",
        ["2"]
    };

    expr_func_call_24 => {
        "program p\nprint *, abs(-3)\nend program p\n",
        ["3"]
    };

    expr_nested_call_25 => {
        "program p\nprint *, max(1, min(2,3))\nend program p\n",
        ["2"]
    };

    expr_kind_conv_26 => {
        "program p\ninteger :: i\nreal :: r=1.5\ni = int(r)\nprint *, i\nend program p\n",
        ["1"]
    };

    expr_real_conv_27 => {
        "program p\nreal :: r\nr = real(3)\nprint *, r\nend program p\n",
        ["3"]
    };

    expr_merge_28 => {
        "program p\ninteger :: x\nx = merge(1,2,.true.)\nprint *, x\nend program p\n",
        ["1"]
    };

    expr_implied_do_29 => {
        "program p\ninteger :: a(3)\na = [(i, i=1,3)]\nprint *, a(1) + a(2) + a(3)\nend program p\n",
        ["6"]
    };

    expr_masked_where_30 => {
        "program p\ninteger :: a(3)=[1,2,3]\nwhere (a > 1) a = a + 1\nprint *, sum(a)\nend program p\n",
        ["8"]
    };

    expr_unary_plus_31 => {
        "program p\ninteger :: a\na = +7\nprint *, a\nend program p\n",
        ["7"]
    };

    expr_nested_parens_32 => {
        "program p\ninteger :: a\na = (10 - 4) * (3 + 1)\nprint *, a\nend program p\n",
        ["24"]
    };

    expr_power_precedence_33 => {
        "program p\ninteger :: a\na = -2 ** 3\nprint *, a\nend program p\n",
        ["-8"]
    };

    expr_power_assoc_34 => {
        "program p\ninteger :: a\na = 2 ** 3 ** 2\nprint *, a\nend program p\n",
        ["512"]
    };

    expr_mixed_type_sub_add_35 => {
        "program p\nprint *, 1 + 2.0 + 3\nend program p\n",
        ["6"]
    };

    expr_char_concat_chain_36 => {
        "program p\ncharacter(len=3) :: s\ns = 'a'//'b'//'c'\nprint *, s\nend program p\n",
        ["abc"]
    };

    expr_logical_group_37 => {
        "program p\nlogical :: x\nx = .true. .and. (.false. .or. .true.)\nprint *, x\nend program p\n",
        ["true"]
    };

    expr_int_division_trunc_38 => {
        "program p\nprint *, 7 / 2\nprint *, -17 / 5\nend program p\n",
        ["3", "-3"]
    };
}
