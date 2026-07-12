use crate::helpers::run_python_one;

#[test]
fn math_sqrt() {
    assert_eq!(run_python_one("import math\nprint(math.sqrt(9))\n"), "3.0");
}

#[test]
fn math_pow() {
    assert_eq!(
        run_python_one("import math\nprint(math.pow(2, 3))\n"),
        "8.0"
    );
}

#[test]
fn math_floor() {
    assert_eq!(run_python_one("import math\nprint(math.floor(3.7))\n"), "3");
}

#[test]
fn math_ceil() {
    assert_eq!(run_python_one("import math\nprint(math.ceil(3.2))\n"), "4");
}

#[test]
fn math_factorial() {
    assert_eq!(
        run_python_one("import math\nprint(math.factorial(5))\n"),
        "120"
    );
}

#[test]
fn math_gcd() {
    assert_eq!(
        run_python_one("import math\nprint(math.gcd(12, 18))\n"),
        "6"
    );
}

#[test]
fn math_lcm() {
    assert_eq!(run_python_one("import math\nprint(math.lcm(4, 6))\n"), "12");
}

#[test]
fn math_hypot() {
    assert_eq!(
        run_python_one("import math\nprint(math.hypot(3, 4))\n"),
        "5.0"
    );
}

#[test]
fn math_sin_zero() {
    assert_eq!(run_python_one("import math\nprint(math.sin(0))\n"), "0.0");
}

#[test]
fn math_cos_zero() {
    assert_eq!(run_python_one("import math\nprint(math.cos(0))\n"), "1.0");
}

#[test]
fn math_tan_zero() {
    assert_eq!(run_python_one("import math\nprint(math.tan(0))\n"), "0.0");
}

#[test]
fn math_log_natural() {
    assert_eq!(
        run_python_one("import math\nprint(round(math.log(math.e), 5))\n"),
        "1.0"
    );
}

#[test]
fn math_log10() {
    assert_eq!(
        run_python_one("import math\nprint(math.log10(1000))\n"),
        "3.0"
    );
}

#[test]
fn math_log2() {
    assert_eq!(run_python_one("import math\nprint(math.log2(8))\n"), "3.0");
}

#[test]
fn math_degrees() {
    assert_eq!(
        run_python_one("import math\nprint(math.degrees(math.pi))\n"),
        "180.0"
    );
}

#[test]
fn math_radians() {
    assert_eq!(
        run_python_one("import math\nprint(math.radians(180))\n"),
        "3.141592653589793"
    );
}

#[test]
fn math_fabs() {
    assert_eq!(
        run_python_one("import math\nprint(math.fabs(-3.5))\n"),
        "3.5"
    );
}

#[test]
fn math_trunc() {
    assert_eq!(run_python_one("import math\nprint(math.trunc(3.9))\n"), "3");
}

#[test]
fn math_isfinite() {
    assert_eq!(
        run_python_one("import math\nprint(math.isfinite(1.0))\n"),
        "True"
    );
}

#[test]
fn math_isinf() {
    assert_eq!(
        run_python_one("import math\nprint(math.isinf(math.inf))\n"),
        "True"
    );
}

#[test]
fn math_isnan() {
    assert_eq!(
        run_python_one("import math\nprint(math.isnan(float('nan')))\n"),
        "True"
    );
}

#[test]
fn math_pi_constant() {
    assert_eq!(
        run_python_one("import math\nprint(round(math.pi, 2))\n"),
        "3.14"
    );
}

#[test]
fn math_e_constant() {
    assert_eq!(
        run_python_one("import math\nprint(round(math.e, 2))\n"),
        "2.72"
    );
}

#[test]
fn math_tau_constant() {
    assert_eq!(
        run_python_one("import math\nprint(round(math.tau, 2))\n"),
        "6.28"
    );
}

#[test]
fn math_copysign() {
    assert_eq!(
        run_python_one("import math\nprint(math.copysign(1, -1))\n"),
        "-1.0"
    );
}

#[test]
fn math_fmod() {
    assert_eq!(
        run_python_one("import math\nprint(math.fmod(7, 3))\n"),
        "1.0"
    );
}

#[test]
fn math_remainder() {
    assert_eq!(
        run_python_one("import math\nprint(math.remainder(7, 3))\n"),
        "1.0"
    );
}

#[test]
fn math_modf() {
    assert_eq!(
        run_python_one("import math\nprint(math.modf(3.75))\n"),
        "(0.75, 3.0)"
    );
}

#[test]
fn math_frexp() {
    assert_eq!(
        run_python_one("import math\nprint(math.frexp(8))\n"),
        "(0.5, 4)"
    );
}

#[test]
fn math_ldexp() {
    assert_eq!(
        run_python_one("import math\nprint(math.ldexp(0.5, 4))\n"),
        "8.0"
    );
}

#[test]
fn math_isclose_true() {
    assert_eq!(
        run_python_one("import math\nprint(math.isclose(0.1 + 0.2, 0.3))\n"),
        "False"
    );
}

#[test]
fn math_isclose_with_tol() {
    assert_eq!(
        run_python_one("import math\nprint(math.isclose(1.0, 1.00001, rel_tol=1e-3))\n"),
        "True"
    );
}

#[test]
fn math_dist_euclidean() {
    assert_eq!(
        run_python_one("import math\nprint(math.dist((0, 0), (3, 4)))\n"),
        "5.0"
    );
}

#[test]
fn math_prod() {
    assert_eq!(
        run_python_one("import math\nprint(math.prod([1, 2, 3, 4]))\n"),
        "24"
    );
}

#[test]
fn math_comb() {
    assert_eq!(
        run_python_one("import math\nprint(math.comb(5, 2))\n"),
        "10"
    );
}

#[test]
fn math_perm() {
    assert_eq!(
        run_python_one("import math\nprint(math.perm(5, 2))\n"),
        "20"
    );
}

#[test]
fn math_ceil_division_via_floor() {
    assert_eq!(
        run_python_one("import math\nprint(-math.floor(-7 / 3))\n"),
        "3"
    );
}

#[test]
fn math_sqrt_of_two() {
    assert_eq!(
        run_python_one("import math\nprint(round(math.sqrt(2), 3))\n"),
        "1.414"
    );
}

#[test]
fn math_exp_one() {
    assert_eq!(
        run_python_one("import math\nprint(round(math.exp(1), 2))\n"),
        "2.72"
    );
}

#[test]
fn math_asin_bound() {
    assert_eq!(run_python_one("import math\nprint(math.asin(0))\n"), "0.0");
}

#[test]
fn math_acos_bound() {
    assert_eq!(run_python_one("import math\nprint(math.acos(1))\n"), "0.0");
}

#[test]
fn math_atan2() {
    assert_eq!(
        run_python_one("import math\nprint(math.atan2(1, 1) > 0)\n"),
        "True"
    );
}

#[test]
fn math_sinh_zero() {
    assert_eq!(run_python_one("import math\nprint(math.sinh(0))\n"), "0.0");
}

#[test]
fn math_cosh_zero() {
    assert_eq!(run_python_one("import math\nprint(math.cosh(0))\n"), "1.0");
}

#[test]
fn math_tanh_zero() {
    assert_eq!(run_python_one("import math\nprint(math.tanh(0))\n"), "0.0");
}

#[test]
fn math_degrees_right_angle_sin() {
    assert_eq!(
        run_python_one("import math\nprint(round(math.sin(math.radians(90)), 5))\n"),
        "1.0"
    );
}
