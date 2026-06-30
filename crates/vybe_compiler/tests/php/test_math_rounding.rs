//! `ceil`, `floor`, `round`, `fmod`, `intdiv`, and rounding mode behavior.

crate::php_cases! {
    ceil_rounds_up_to_integer => {
        r#"<?php
echo ceil(2.1);
"#,
        ["3"]
    };

    floor_rounds_down_to_integer => {
        r#"<?php
echo floor(2.9);
"#,
        ["2"]
    };

    round_half_up_default => {
        r#"<?php
echo round(2.5);
"#,
        ["3"]
    };

    round_with_precision_two_decimals => {
        r#"<?php
echo round(3.14159, 2);
"#,
        ["3.14"]
    };

    fmod_returns_float_remainder => {
        r#"<?php
echo fmod(5.7, 1.3);
"#,
        ["0.5"]
    };

    intdiv_truncates_toward_zero => {
        r#"<?php
echo intdiv(7, 2);
"#,
        ["3"]
    };

    intdiv_negative_truncates_toward_zero => {
        r#"<?php
echo intdiv(-7, 2);
"#,
        ["-3"]
    };

    fdiv_float_division => {
        r#"<?php
echo fdiv(5.0, 2.0);
"#,
        ["2.5"]
    };

    hypot_three_four_five_triangle => {
        r#"<?php
echo hypot(3, 4);
"#,
        ["5"]
    };

    is_nan_on_nan_value => {
        r#"<?php
echo is_nan(NAN) ? 'nan' : 'num';
"#,
        ["nan"]
    };

    is_finite_rejects_infinity => {
        r#"<?php
echo is_finite(INF) ? 'fin' : 'inf';
"#,
        ["inf"]
    };

    deg2rad_180_is_pi => {
        r#"<?php
echo round(deg2rad(180), 5);
"#,
        ["3.14159"]
    };

    rad2deg_pi_is_180 => {
        r#"<?php
echo round(rad2deg(M_PI));
"#,
        ["180"]
    };

    lcg_value_between_zero_and_one => {
        r#"<?php
$v = lcg_value();
echo ($v >= 0 && $v < 1) ? 'range' : 'out';
"#,
        ["range"]
    };
}
