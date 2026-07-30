use super::helpers::{compile_ok, run_prints};

// ── Integer kinds ─────────────────────────────────────────────

#[test]
fn integer_kind_4() {
    compile_ok("program t\n  integer(kind=4) :: x = 100\n  print *, x\nend program t\n");
}

#[test]
fn integer_kind_8() {
    compile_ok("program t\n  integer(kind=8) :: big = 1000000000\n  print *, big\nend program t\n");
}

#[test]
fn integer_kind_2() {
    compile_ok("program t\n  integer(kind=2) :: s = 32000\n  print *, s\nend program t\n");
}

#[test]
fn integer_kind_1() {
    compile_ok("program t\n  integer(kind=1) :: b = 127\n  print *, b\nend program t\n");
}

#[test]
fn int32_param() {
    compile_ok(
        r#"
program test
    integer, parameter :: int32 = 4
    integer(kind=int32) :: x = 2147483647
    print *, x
end program test
"#,
    );
}

#[test]
fn int64_param() {
    compile_ok(
        r#"
program test
    integer, parameter :: int64 = 8
    integer(kind=int64) :: big = 100000000000_8
    print *, big
end program test
"#,
    );
}

// ── Real kinds ───────────────────────────────────────────────

#[test]
fn real_kind_4() {
    compile_ok("program t\n  real(kind=4) :: x = 3.14\n  print *, x\nend program t\n");
}

#[test]
fn real_kind_8() {
    compile_ok("program t\n  real(kind=8) :: d = 3.14159265358979\n  print *, d\nend program t\n");
}

#[test]
fn real_kind_16() {
    compile_ok(
        "program t\n  real(kind=16) :: q = 3.14159265358979_16\n  print *, q\nend program t\n",
    );
}

#[test]
fn real64_param() {
    compile_ok(
        r#"
program test
    integer, parameter :: real64 = 8
    real(kind=real64) :: x = 1.0_8
    print *, x
end program test
"#,
    );
}

// ── Double precision ─────────────────────────────────────────

#[test]
fn double_precision_assign() {
    compile_ok(
        "program t\n  double precision :: d\n  d = 3.141592653589793\n  print *, d\nend program t\n",
    );
}

#[test]
fn double_precision_arithmetic() {
    compile_ok(
        "program t\n  double precision :: a = 1.0, b = 3.0, c\n  c = a / b\n  print *, c\nend program t\n",
    );
}

#[test]
fn double_precision_parameter() {
    compile_ok(
        "program t\n  double precision, parameter :: PI = 3.141592653589793\n  print *, PI\nend program t\n",
    );
}

// ── Kind literals ─────────────────────────────────────────────

#[test]
fn int_literal_kind() {
    compile_ok("program t\n  integer :: x = 100_4\n  print *, x\nend program t\n");
}

#[test]
fn real_literal_kind() {
    compile_ok("program t\n  real :: x = 3.14_4\n  print *, x\nend program t\n");
}

#[test]
fn double_literal() {
    compile_ok("program t\n  real(kind=8) :: d = 1.0d0\n  print *, d\nend program t\n");
}

#[test]
fn double_literal_exponent() {
    compile_ok("program t\n  real(kind=8) :: d = 1.23d+10\n  print *, d\nend program t\n");
}

#[test]
fn kind_intrinsic_double_literal_runtime() {
    let out = run_prints(
        "program t\n  integer, parameter :: dp = kind(1.0d0)\n  print *, dp\nend program t\n",
    );
    assert_eq!(out, vec!["8"]);
}

// ── Selected kind queries ─────────────────────────────────────

#[test]
fn selected_int_kind_9() {
    compile_ok(
        r#"
program test
    integer, parameter :: k = selected_int_kind(9)
    integer(kind=k) :: n = 999999999
    print *, n
end program test
"#,
    );
}

#[test]
fn selected_real_kind_15() {
    compile_ok(
        r#"
program test
    integer, parameter :: k = selected_real_kind(15, 307)
    real(kind=k) :: x = 1.23456789012345_k
    print *, x
end program test
"#,
    );
}

#[test]
fn selected_int_kind_and_kind_runtime() {
    let out = run_prints(
        r#"
program test
    print *, kind(1)
    print *, selected_int_kind(9)
    print *, selected_real_kind(15, 307)
end program test
"#,
    );
    assert_eq!(out, ["8", "8", "8"]);
}

#[test]
fn kind_queries_for_multiple_types() {
    let out = run_prints(
        r#"
program test
    logical :: l
    complex :: c
    character(len=5) :: s
    integer(kind=4) :: i
    real(kind=8) :: r
    l = .true.
    c = (1.0, 2.0)
    s = "abc"
    i = 5
    r = 3.0
    print *, kind(l)
    print *, kind(c)
    print *, kind(s)
    print *, kind(i)
    print *, kind(r)
end program test
"#,
    );
    assert_eq!(out, ["8", "8", "8", "4", "8"]);
}

#[test]
fn kind_aliases_with_arrays_and_procedure() {
    let out = run_prints(
        r#"
program test
    integer, parameter :: ki = selected_int_kind(9)
    real, parameter :: kr = selected_real_kind(6, 37)
    integer(kind=ki), dimension(4) :: a = [1, 2, 3, 4]
    real(kind=kr), dimension(3) :: b = [1.0, 2.0, 3.0]
    print *, size(a)
    print *, size(b)
    print *, kind(a)
    print *, kind(b)
    print *, sum(a)
end program test
"#,
    );
    assert_eq!(out, ["4", "3", "8", "8", "10"]);
}

// ── Kind in derived types ─────────────────────────────────────

#[test]
fn kind_in_derived_type() {
    compile_ok(
        r#"
program test
    integer, parameter :: wp = 8
    type :: HighPrec
        real(kind=wp) :: value
    end type HighPrec
    type(HighPrec) :: h
    h%value = 1.23456789012345_wp
    print *, h%value
end program test
"#,
    );
}

// ── Kind in function signatures ───────────────────────────────

#[test]
fn kind_in_function() {
    compile_ok(
        r#"
program test
    integer, parameter :: dp = 8
    print *, dp_add(1.0_dp, 2.0_dp)
contains
    function dp_add(a, b) result(res)
        integer, parameter :: dp = 8
        real(kind=dp), intent(in) :: a, b
        real(kind=dp) :: res
        res = a + b
    end function dp_add
end program test
"#,
    );
}

// ── Kind with arrays ──────────────────────────────────────────

#[test]
fn kind_array() {
    compile_ok(
        r#"
program test
    integer(kind=8) :: a(3) = [1_8, 2_8, 3_8]
    print *, a(1)
end program test
"#,
    );
}

#[test]
fn kind_real_array() {
    compile_ok(
        r#"
program test
    real(kind=8) :: v(3) = [1.0_8, 2.0_8, 3.0_8]
    print *, v(1)
end program test
"#,
    );
}
