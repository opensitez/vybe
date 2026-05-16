use super::helpers::compile_ok;

// ── NORM2 ─────────────────────────────────────────────────────

#[test] fn norm2_basic() {
    compile_ok(r#"
program test
    real :: a(3) = [3.0, 4.0, 0.0]
    print *, norm2(a)
end program test
"#);
}

#[test] fn norm2_unit_vector() {
    compile_ok(r#"
program test
    real :: a(3) = [1.0, 0.0, 0.0]
    print *, norm2(a)
end program test
"#);
}

#[test] fn norm2_2d_dim1() {
    compile_ok(r#"
program test
    real :: m(2,3) = reshape([1.,2.,3.,4.,5.,6.],[2,3])
    real :: col_norms(3)
    col_norms = norm2(m, dim=1)
    print *, col_norms(1)
end program test
"#);
}

#[test] fn norm2_double() {
    compile_ok(r#"
program test
    real(kind=8) :: v(4) = [1.0d0, 1.0d0, 1.0d0, 1.0d0]
    print *, norm2(v)
end program test
"#);
}

// ── HYPOT ─────────────────────────────────────────────────────

#[test] fn hypot_3_4_5() {
    compile_ok("program t\n  print *, hypot(3.0, 4.0)\nend program t\n");
}

#[test] fn hypot_zero_y() {
    compile_ok("program t\n  print *, hypot(5.0, 0.0)\nend program t\n");
}

#[test] fn hypot_zero_x() {
    compile_ok("program t\n  print *, hypot(0.0, 5.0)\nend program t\n");
}

#[test] fn hypot_double() {
    compile_ok("program t\n  print *, hypot(3.0d0, 4.0d0)\nend program t\n");
}

#[test] fn hypot_unit_circle() {
    compile_ok(r#"
program test
    real, parameter :: pi = 3.14159265
    real :: angle = pi / 4.0
    real :: h
    h = hypot(cos(angle), sin(angle))
    print *, h
end program test
"#);
}

// ── ACOSH / ASINH / ATANH ─────────────────────────────────────

#[test] fn acosh_one() {
    compile_ok("program t\n  print *, acosh(1.0)\nend program t\n");
}

#[test] fn acosh_cosh_roundtrip() {
    compile_ok(r#"
program test
    real :: x = 2.0
    print *, acosh(cosh(x))
end program test
"#);
}

#[test] fn asinh_zero() {
    compile_ok("program t\n  print *, asinh(0.0)\nend program t\n");
}

#[test] fn asinh_sinh_roundtrip() {
    compile_ok(r#"
program test
    real :: x = 1.5
    print *, asinh(sinh(x))
end program test
"#);
}

#[test] fn atanh_zero() {
    compile_ok("program t\n  print *, atanh(0.0)\nend program t\n");
}

#[test] fn atanh_tanh_roundtrip() {
    compile_ok(r#"
program test
    real :: x = 0.5
    print *, atanh(tanh(x))
end program test
"#);
}

#[test] fn inverse_hyp_double() {
    compile_ok(r#"
program test
    real(kind=8) :: x = 1.0d0
    print *, acosh(cosh(x))
    print *, asinh(sinh(x))
    print *, atanh(tanh(x * 0.5d0))
end program test
"#);
}

// ── ERF / ERFC / ERFC_SCALED ──────────────────────────────────

#[test] fn erf_zero() {
    compile_ok("program t\n  print *, erf(0.0)\nend program t\n");
}

#[test] fn erf_positive() {
    compile_ok("program t\n  print *, erf(1.0)\nend program t\n");
}

#[test] fn erf_large() {
    compile_ok("program t\n  print *, erf(10.0)\nend program t\n");
}

#[test] fn erf_negative() {
    compile_ok("program t\n  print *, erf(-1.0)\nend program t\n");
}

#[test] fn erfc_zero() {
    compile_ok("program t\n  print *, erfc(0.0)\nend program t\n");
}

#[test] fn erfc_one() {
    compile_ok("program t\n  print *, erfc(1.0)\nend program t\n");
}

#[test] fn erf_erfc_sum_is_one() {
    compile_ok(r#"
program test
    real :: x = 1.5
    real :: total
    total = erf(x) + erfc(x)
    print *, total
end program test
"#);
}

#[test] fn erfc_scaled_basic() {
    compile_ok("program t\n  print *, erfc_scaled(1.0)\nend program t\n");
}

#[test] fn erfc_scaled_large() {
    compile_ok("program t\n  print *, erfc_scaled(10.0)\nend program t\n");
}

#[test] fn erf_double() {
    compile_ok("program t\n  print *, erf(1.0d0)\nend program t\n");
}

// ── GAMMA / LOG_GAMMA ─────────────────────────────────────────

#[test] fn gamma_one() {
    compile_ok("program t\n  print *, gamma(1.0)\nend program t\n");
}

#[test] fn gamma_two() {
    compile_ok("program t\n  print *, gamma(2.0)\nend program t\n");
}

#[test] fn gamma_five() {
    compile_ok("program t\n  print *, gamma(5.0)\nend program t\n");
}

#[test] fn gamma_half() {
    compile_ok("program t\n  print *, gamma(0.5)\nend program t\n");
}

#[test] fn gamma_double() {
    compile_ok("program t\n  print *, gamma(1.0d0)\nend program t\n");
}

#[test] fn log_gamma_one() {
    compile_ok("program t\n  print *, log_gamma(1.0)\nend program t\n");
}

#[test] fn log_gamma_positive() {
    compile_ok("program t\n  print *, log_gamma(5.0)\nend program t\n");
}

#[test] fn log_gamma_vs_log_gamma() {
    compile_ok(r#"
program test
    real :: x = 10.0
    print *, log_gamma(x)
    print *, log(gamma(x))
end program test
"#);
}

// ── BESSEL_J0 / BESSEL_J1 / BESSEL_JN ────────────────────────

#[test] fn bessel_j0_zero() {
    compile_ok("program t\n  print *, bessel_j0(0.0)\nend program t\n");
}

#[test] fn bessel_j0_positive() {
    compile_ok("program t\n  print *, bessel_j0(1.0)\nend program t\n");
}

#[test] fn bessel_j1_zero() {
    compile_ok("program t\n  print *, bessel_j1(0.0)\nend program t\n");
}

#[test] fn bessel_j1_positive() {
    compile_ok("program t\n  print *, bessel_j1(1.0)\nend program t\n");
}

#[test] fn bessel_jn_order_0() {
    compile_ok("program t\n  print *, bessel_jn(0, 1.0)\nend program t\n");
}

#[test] fn bessel_jn_order_2() {
    compile_ok("program t\n  print *, bessel_jn(2, 1.0)\nend program t\n");
}

#[test] fn bessel_jn_array() {
    compile_ok(r#"
program test
    real :: values(3)
    values = bessel_jn(0, 2, 1.0)
    print *, values(1)
end program test
"#);
}

#[test] fn bessel_j0_double() {
    compile_ok("program t\n  print *, bessel_j0(1.0d0)\nend program t\n");
}

// ── BESSEL_Y0 / BESSEL_Y1 / BESSEL_YN ────────────────────────

#[test] fn bessel_y0_positive() {
    compile_ok("program t\n  print *, bessel_y0(1.0)\nend program t\n");
}

#[test] fn bessel_y1_positive() {
    compile_ok("program t\n  print *, bessel_y1(1.0)\nend program t\n");
}

#[test] fn bessel_yn_order_0() {
    compile_ok("program t\n  print *, bessel_yn(0, 1.0)\nend program t\n");
}

#[test] fn bessel_yn_order_2() {
    compile_ok("program t\n  print *, bessel_yn(2, 1.0)\nend program t\n");
}

#[test] fn bessel_yn_array() {
    compile_ok(r#"
program test
    real :: values(3)
    values = bessel_yn(0, 2, 1.0)
    print *, values(1)
end program test
"#);
}

// ── Combining F2008 math intrinsics ───────────────────────────

#[test] fn combined_norm2_hypot() {
    compile_ok(r#"
program test
    real :: v(2) = [3.0, 4.0]
    real :: h
    h = hypot(v(1), v(2))
    print *, norm2(v)
    print *, h
end program test
"#);
}

#[test] fn erf_with_gamma() {
    compile_ok(r#"
program test
    real :: x = 0.5
    print *, erf(x)
    print *, gamma(x + 0.5) / gamma(0.5)
end program test
"#);
}
