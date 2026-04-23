use super::helpers::compile_ok;

// ── COMPLEX declarations ──────────────────────────────────────

#[test] fn complex_literal() { compile_ok("program t\n  complex :: z = (3.0, 4.0)\n  print *, z\nend program t\n"); }
#[test] fn complex_zero() { compile_ok("program t\n  complex :: z = (0.0, 0.0)\n  print *, z\nend program t\n"); }
#[test] fn complex_pure_imaginary() { compile_ok("program t\n  complex :: z = (0.0, 1.0)\n  print *, z\nend program t\n"); }
#[test] fn complex_negative() { compile_ok("program t\n  complex :: z = (-1.0, -2.0)\n  print *, z\nend program t\n"); }
#[test] fn complex_double() { compile_ok("program t\n  double complex :: z = (1.0d0, 2.0d0)\n  print *, z\nend program t\n"); }
#[test] fn complex_kind8() { compile_ok("program t\n  complex(kind=8) :: z = (1.0_8, 2.0_8)\n  print *, z\nend program t\n"); }

// ── CMPLX constructor ─────────────────────────────────────────

#[test] fn cmplx_two_args() { compile_ok("program t\n  complex :: z\n  z = cmplx(3.0, 4.0)\n  print *, z\nend program t\n"); }
#[test] fn cmplx_one_arg() { compile_ok("program t\n  complex :: z\n  z = cmplx(5.0)\n  print *, z\nend program t\n"); }
#[test] fn cmplx_from_int() { compile_ok("program t\n  complex :: z\n  integer :: a = 3, b = 4\n  z = cmplx(a, b)\n  print *, z\nend program t\n"); }
#[test] fn cmplx_kind() { compile_ok("program t\n  complex(kind=8) :: z\n  z = cmplx(1.0, 2.0, kind=8)\n  print *, z\nend program t\n"); }

// ── REAL and AIMAG ─────────────────────────────────────────────

#[test] fn real_part() { compile_ok("program t\n  complex :: z = (3.0, 4.0)\n  print *, real(z)\nend program t\n"); }
#[test] fn aimag_part() { compile_ok("program t\n  complex :: z = (3.0, 4.0)\n  print *, aimag(z)\nend program t\n"); }
#[test] fn real_aimag_both() {
    compile_ok(r#"
program test
    complex :: z = (3.0, 4.0)
    real :: r, i
    r = real(z)
    i = aimag(z)
    print *, r
    print *, i
end program test
"#);
}

// ── CONJG ────────────────────────────────────────────────────

#[test] fn conjg_basic() { compile_ok("program t\n  complex :: z = (3.0, 4.0)\n  complex :: c\n  c = conjg(z)\n  print *, c\nend program t\n"); }
#[test] fn conjg_pure_real() { compile_ok("program t\n  complex :: z = (5.0, 0.0)\n  print *, conjg(z)\nend program t\n"); }

// ── ABS of complex ─────────────────────────────────────────────

#[test] fn abs_complex() { compile_ok("program t\n  complex :: z = (3.0, 4.0)\n  real :: m\n  m = abs(z)\n  print *, m\nend program t\n"); }
#[test] fn abs_unit_imaginary() { compile_ok("program t\n  complex :: z = (0.0, 1.0)\n  print *, abs(z)\nend program t\n"); }

// ── Complex arithmetic ────────────────────────────────────────

#[test] fn complex_add() {
    compile_ok(r#"
program test
    complex :: a = (1.0, 2.0), b = (3.0, 4.0), c
    c = a + b
    print *, c
end program test
"#);
}

#[test] fn complex_sub() {
    compile_ok(r#"
program test
    complex :: a = (5.0, 6.0), b = (2.0, 3.0), c
    c = a - b
    print *, c
end program test
"#);
}

#[test] fn complex_mul() {
    compile_ok(r#"
program test
    complex :: a = (1.0, 2.0), b = (3.0, 4.0), c
    c = a * b
    print *, c
end program test
"#);
}

#[test] fn complex_div() {
    compile_ok(r#"
program test
    complex :: a = (1.0, 0.0), b = (2.0, 0.0), c
    c = a / b
    print *, c
end program test
"#);
}

#[test] fn complex_power() {
    compile_ok(r#"
program test
    complex :: z = (1.0, 1.0)
    complex :: r
    r = z ** 2
    print *, r
end program test
"#);
}

#[test] fn complex_negate() {
    compile_ok(r#"
program test
    complex :: z = (3.0, 4.0)
    complex :: n
    n = -z
    print *, n
end program test
"#);
}

#[test] fn complex_mixed_real() {
    compile_ok(r#"
program test
    complex :: z = (3.0, 4.0)
    complex :: r
    r = 2.0 * z
    print *, r
end program test
"#);
}

// ── Complex intrinsics ────────────────────────────────────────

#[test] fn sqrt_complex() { compile_ok("program t\n  complex :: z = (-1.0, 0.0)\n  complex :: r\n  r = sqrt(z)\n  print *, r\nend program t\n"); }
#[test] fn exp_complex() { compile_ok("program t\n  complex :: z = (0.0, 3.14159)\n  complex :: r\n  r = exp(z)\n  print *, r\nend program t\n"); }
#[test] fn log_complex() { compile_ok("program t\n  complex :: z = (1.0, 0.0)\n  complex :: r\n  r = log(z)\n  print *, r\nend program t\n"); }
#[test] fn sin_complex() { compile_ok("program t\n  complex :: z = (0.0, 0.0)\n  complex :: r\n  r = sin(z)\n  print *, r\nend program t\n"); }
#[test] fn cos_complex() { compile_ok("program t\n  complex :: z = (0.0, 0.0)\n  complex :: r\n  r = cos(z)\n  print *, r\nend program t\n"); }

// ── Complex in arrays ─────────────────────────────────────────

#[test] fn complex_array() {
    compile_ok(r#"
program test
    complex :: v(3) = [(0.0,0.0), (1.0,0.0), (0.0,1.0)]
    print *, real(v(2))
end program test
"#);
}

#[test] fn complex_array_ops() {
    compile_ok(r#"
program test
    complex :: a(3) = [(1.0,0.0), (2.0,0.0), (3.0,0.0)]
    complex :: b(3) = [(0.0,1.0), (0.0,2.0), (0.0,3.0)]
    complex :: c(3)
    c = a + b
    print *, real(c(1))
end program test
"#);
}

// ── Complex in derived types ──────────────────────────────────

#[test] fn complex_in_type() {
    compile_ok(r#"
program test
    type :: Phasor
        real :: magnitude
        complex :: value
    end type Phasor
    type(Phasor) :: p
    p%magnitude = 5.0
    p%value = (3.0, 4.0)
    print *, p%magnitude
end program test
"#);
}

// ── Complex comparison ────────────────────────────────────────

#[test] fn complex_equal() {
    compile_ok(r#"
program test
    complex :: a = (1.0, 2.0), b = (1.0, 2.0)
    print *, a == b
end program test
"#);
}

#[test] fn complex_not_equal() {
    compile_ok(r#"
program test
    complex :: a = (1.0, 2.0), b = (1.0, 3.0)
    print *, a /= b
end program test
"#);
}
