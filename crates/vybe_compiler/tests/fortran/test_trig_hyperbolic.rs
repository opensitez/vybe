//! Fortran trigonometric and hyperbolic intrinsics with known-value checks.

fortran_cases! {
    sin_at_zero_scaled => {
        "program t\nprint *, nint(sin(0.0)*100)\nend program t\n",
        ["0"]
    };

    cos_at_zero_scaled => {
        "program t\nprint *, nint(cos(0.0)*100)\nend program t\n",
        ["100"]
    };

    tan_at_zero_scaled => {
        "program t\nprint *, nint(tan(0.0)*100)\nend program t\n",
        ["0"]
    };

    sin_pi_over_six => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(sin(pi/6.0)*100)\nend program t\n",
        ["50"]
    };

    sin_pi_over_four => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(sin(pi/4.0)*100)\nend program t\n",
        ["71"]
    };

    sin_pi_over_three => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(sin(pi/3.0)*100)\nend program t\n",
        ["87"]
    };

    sin_pi_over_two => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(sin(pi/2.0)*100)\nend program t\n",
        ["100"]
    };

    sin_pi => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(sin(pi)*100)\nend program t\n",
        ["0"]
    };

    sin_neg_pi_over_four => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(sin(-pi/4.0)*100)\nend program t\n",
        ["-71"]
    };

    sin_three_pi_over_two => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(sin(3.0*pi/2.0)*100)\nend program t\n",
        ["-100"]
    };

    cos_pi_over_six => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(cos(pi/6.0)*100)\nend program t\n",
        ["87"]
    };

    cos_pi_over_four => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(cos(pi/4.0)*100)\nend program t\n",
        ["71"]
    };

    cos_pi_over_three => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(cos(pi/3.0)*100)\nend program t\n",
        ["50"]
    };

    cos_pi_over_two => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(cos(pi/2.0)*100)\nend program t\n",
        ["0"]
    };

    cos_pi => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(cos(pi)*100)\nend program t\n",
        ["-100"]
    };

    cos_two_pi => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(cos(2.0*pi)*100)\nend program t\n",
        ["100"]
    };

    tan_pi_over_six => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(tan(pi/6.0)*100)\nend program t\n",
        ["58"]
    };

    tan_pi_over_four => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(tan(pi/4.0)*100)\nend program t\n",
        ["100"]
    };

    tan_pi_over_three => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(tan(pi/3.0)*100)\nend program t\n",
        ["173"]
    };

    tan_neg_pi_over_four => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(tan(-pi/4.0)*100)\nend program t\n",
        ["-100"]
    };

    asin_zero_scaled => {
        "program t\nprint *, nint(asin(0.0)*1000)\nend program t\n",
        ["0"]
    };

    asin_one_half => {
        "program t\nprint *, nint(asin(0.5)*1000)\nend program t\n",
        ["524"]
    };

    asin_neg_one_half => {
        "program t\nprint *, nint(asin(-0.5)*1000)\nend program t\n",
        ["-524"]
    };

    asin_one => {
        "program t\nprint *, nint(asin(1.0)*1000)\nend program t\n",
        ["1571"]
    };

    acos_one_scaled => {
        "program t\nprint *, nint(acos(1.0)*1000)\nend program t\n",
        ["0"]
    };

    acos_zero => {
        "program t\nprint *, nint(acos(0.0)*1000)\nend program t\n",
        ["1571"]
    };

    acos_one_half => {
        "program t\nprint *, nint(acos(0.5)*1000)\nend program t\n",
        ["1047"]
    };

    acos_neg_one => {
        "program t\nprint *, nint(acos(-1.0)*1000)\nend program t\n",
        ["3142"]
    };

    atan_one => {
        "program t\nprint *, nint(atan(1.0)*1000)\nend program t\n",
        ["785"]
    };

    atan_zero_scaled => {
        "program t\nprint *, nint(atan(0.0)*1000)\nend program t\n",
        ["0"]
    };

    atan_sqrt_three => {
        "program t\nprint *, nint(atan(sqrt(3.0))*1000)\nend program t\n",
        ["1047"]
    };

    atan2_one_one => {
        "program t\nprint *, nint(atan2(1.0, 1.0)*1000)\nend program t\n",
        ["785"]
    };

    atan2_neg_one_one => {
        "program t\nprint *, nint(atan2(-1.0, 1.0)*1000)\nend program t\n",
        ["2356"]
    };

    atan2_one_neg_one => {
        "program t\nprint *, nint(atan2(1.0, -1.0)*1000)\nend program t\n",
        ["-785"]
    };

    atan2_neg_one_neg_one => {
        "program t\nprint *, nint(atan2(-1.0, -1.0)*1000)\nend program t\n",
        ["-2356"]
    };

    atan2_zero_one => {
        "program t\nprint *, nint(atan2(0.0, 1.0)*1000)\nend program t\n",
        ["0"]
    };

    atan2_one_zero => {
        "program t\nprint *, nint(atan2(1.0, 0.0)*1000)\nend program t\n",
        ["1571"]
    };

    atan2_neg_one_zero => {
        "program t\nprint *, nint(atan2(-1.0, 0.0)*1000)\nend program t\n",
        ["-1571"]
    };

    atan2_zero_neg_one => {
        "program t\nprint *, nint(atan2(0.0, -1.0)*1000)\nend program t\n",
        ["3142"]
    };

    atan2_sqrt_three_one => {
        "program t\nprint *, nint(atan2(sqrt(3.0), 1.0)*1000)\nend program t\n",
        ["1047"]
    };

    sinh_zero_scaled => {
        "program t\nprint *, nint(sinh(0.0)*100)\nend program t\n",
        ["0"]
    };

    cosh_zero_scaled => {
        "program t\nprint *, nint(cosh(0.0)*100)\nend program t\n",
        ["100"]
    };

    tanh_zero_scaled => {
        "program t\nprint *, nint(tanh(0.0)*100)\nend program t\n",
        ["0"]
    };

    sinh_small_one_tenth => {
        "program t\nprint *, nint(sinh(0.1)*1000)\nend program t\n",
        ["100"]
    };

    sinh_small_one_half => {
        "program t\nprint *, nint(sinh(0.5)*1000)\nend program t\n",
        ["521"]
    };

    cosh_small_one_tenth => {
        "program t\nprint *, nint(cosh(0.1)*1000)\nend program t\n",
        ["1005"]
    };

    cosh_small_one_half => {
        "program t\nprint *, nint(cosh(0.5)*1000)\nend program t\n",
        ["1128"]
    };

    tanh_small_one_half => {
        "program t\nprint *, nint(tanh(0.5)*1000)\nend program t\n",
        ["462"]
    };

    tanh_one => {
        "program t\nprint *, nint(tanh(1.0)*1000)\nend program t\n",
        ["762"]
    };

    sin_cos_pythagorean_pi_over_four => {
        "program t\nreal, parameter :: pi = acos(-1.0)\nreal :: s, c\ns = sin(pi/4.0)\nc = cos(pi/4.0)\nprint *, nint((s*s + c*c)*100)\nend program t\n",
        ["100"]
    };
}
