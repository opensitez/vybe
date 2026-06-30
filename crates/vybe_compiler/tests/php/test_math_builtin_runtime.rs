//! `abs`, trig, log, and hyperbolic math builtins (runtime).

crate::php_cases! {
    abs_negative_integer => {
        r#"<?php
echo abs(-9);
"#,
        ["9"]
    };

    abs_float => {
        r#"<?php
echo (int)abs(-2.5);
"#,
        ["2"]
    };

    ceil_rounds_up => {
        r#"<?php
echo ceil(2.1);
"#,
        ["3"]
    };

    floor_rounds_down => {
        r#"<?php
echo floor(2.9);
"#,
        ["2"]
    };

    round_half_up => {
        r#"<?php
echo round(2.5);
"#,
        ["3"]
    };

    round_precision_two => {
        r#"<?php
echo round(1.234, 2);
"#,
        ["1.23"]
    };

    sqrt_perfect_square => {
        r#"<?php
echo sqrt(16);
"#,
        ["4"]
    };

    pow_fractional => {
        r#"<?php
echo (int)pow(2, 3);
"#,
        ["8"]
    };

    fmod_float_remainder => {
        r#"<?php
echo fmod(5.5, 2);
"#,
        ["1.5"]
    };

    hypot_pythagorean => {
        r#"<?php
echo hypot(3, 4);
"#,
        ["5"]
    };

    sin_zero => {
        r#"<?php
echo sin(0);
"#,
        ["0"]
    };

    cos_zero => {
        r#"<?php
echo cos(0);
"#,
        ["1"]
    };

    tan_zero => {
        r#"<?php
echo tan(0);
"#,
        ["0"]
    };

    deg2rad_180 => {
        r#"<?php
echo (int)round(deg2rad(180) / M_PI);
"#,
        ["1"]
    };

    rad2deg_pi => {
        r#"<?php
echo (int)rad2deg(M_PI);
"#,
        ["180"]
    };

    log_natural_e => {
        r#"<?php
echo (int)round(log(M_E));
"#,
        ["1"]
    };

    log10_thousand => {
        r#"<?php
echo log10(1000);
"#,
        ["3"]
    };

    exp_one => {
        r#"<?php
echo (int)round(exp(0));
"#,
        ["1"]
    };

    min_variadic => {
        r#"<?php
echo min(3, 1, 4, 2);
"#,
        ["1"]
    };

    max_variadic => {
        r#"<?php
echo max(3, 1, 4, 2);
"#,
        ["4"]
    };

    fdiv_float_division => {
        r#"<?php
echo fdiv(5, 2);
"#,
        ["2.5"]
    };

    is_finite_normal => {
        r#"<?php
echo is_finite(1.5) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_infinite_overflow => {
        r#"<?php
echo is_infinite(INF) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_nan_not_number => {
        r#"<?php
echo is_nan(NAN) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    pi_constant => {
        r#"<?php
echo (int)round(M_PI);
"#,
        ["3"]
    };
}
