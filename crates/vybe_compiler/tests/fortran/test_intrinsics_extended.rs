use super::helpers::{compile_ok, run_prints};

// ── Rounding / truncation ─────────────────────────────────────

#[test]
fn nint_round_up() {
    let out = run_prints("program t\nprint *, nint(3.7)\nend program t\n");
    assert_eq!(out, ["4"]);
}

#[test]
fn nint_round_down() {
    let out = run_prints("program t\nprint *, nint(3.2)\nend program t\n");
    assert_eq!(out, ["3"]);
}

#[test]
fn aint_truncate() {
    let out = run_prints("program t\nreal :: x\nx = aint(3.9)\nprint *, x\nend program t\n");
    assert_eq!(out, ["3"]);
}

#[test]
fn anint_nearest() {
    let out = run_prints("program t\nreal :: x\nx = anint(3.5)\nprint *, x\nend program t\n");
    assert_eq!(out, ["4"]);
}

#[test]
fn int_convert() {
    let out = run_prints("program t\nprint *, int(3.9)\nend program t\n");
    assert_eq!(out, ["3"]);
}

#[test]
fn real_convert() {
    let out = run_prints("program t\nprint *, real(7)\nend program t\n");
    assert_eq!(out, ["7"]);
}

#[test]
fn real_convert_with_kind_arg() {
    let out = run_prints(
        "program t\ninteger, parameter :: dp = kind(1.0d0)\nprint *, real(7, dp)\nend program t\n",
    );
    assert_eq!(out, ["7"]);
}

#[test]
fn real_convert_with_kind_arg_in_internal_function() {
    let out = run_prints(
        "program t\ninteger, parameter :: dp = kind(1.0d0)\nprint *, sample(7)\ncontains\npure function sample(s) result(r)\ninteger, intent(in) :: s\nreal(dp) :: r\nr = real(s, dp)\nend function sample\nend program t\n",
    );
    assert_eq!(out, ["7"]);
}

#[test]
fn dble_convert() {
    compile_ok("program t\ndouble precision :: d\nd = dble(3)\nprint *, d\nend program t\n");
}

// ── SIGN ──────────────────────────────────────────────────────

#[test]
fn sign_pos_to_neg() {
    let out = run_prints("program t\nprint *, sign(5, -1)\nend program t\n");
    assert_eq!(out, ["-5"]);
}

#[test]
fn sign_neg_to_pos() {
    let out = run_prints("program t\nprint *, sign(-5, 1)\nend program t\n");
    assert_eq!(out, ["5"]);
}

#[test]
fn sign_real() {
    compile_ok("program t\nreal :: x\nx = sign(3.14, -1.0)\nprint *, x\nend program t\n");
}

// ── DIM ───────────────────────────────────────────────────────

#[test]
fn dim_positive() {
    let out = run_prints("program t\nprint *, dim(10, 3)\nend program t\n");
    assert_eq!(out, ["7"]);
}

#[test]
fn dim_zero() {
    let out = run_prints("program t\nprint *, dim(3, 10)\nend program t\n");
    assert_eq!(out, ["0"]);
}

// ── MODULO ───────────────────────────────────────────────────

#[test]
fn modulo_positive() {
    let out = run_prints("program t\nprint *, modulo(10, 3)\nend program t\n");
    assert_eq!(out, ["1"]);
}

#[test]
fn modulo_negative() {
    let out = run_prints("program t\nprint *, modulo(-1, 5)\nend program t\n");
    assert_eq!(out, ["4"]);
}

// ── Trig — more functions ────────────────────────────────────

#[test]
fn asin_zero() {
    let out = run_prints("program t\nreal :: x\nx = asin(0.0)\nprint *, x\nend program t\n");
    assert_eq!(out, ["0"]);
}

#[test]
fn acos_one() {
    let out = run_prints("program t\nreal :: x\nx = acos(1.0)\nprint *, x\nend program t\n");
    assert_eq!(out, ["0"]);
}

#[test]
fn atan_zero() {
    let out = run_prints("program t\nreal :: x\nx = atan(0.0)\nprint *, x\nend program t\n");
    assert_eq!(out, ["0"]);
}

#[test]
fn atan2_basic() {
    compile_ok("program t\nreal :: x\nx = atan2(1.0, 1.0)\nprint *, x\nend program t\n");
}

#[test]
fn tan_zero() {
    let out = run_prints("program t\nreal :: x\nx = tan(0.0)\nprint *, x\nend program t\n");
    assert_eq!(out, ["0"]);
}

// ── Hyperbolic ────────────────────────────────────────────────

#[test]
fn sinh_zero() {
    let out = run_prints("program t\nreal :: x\nx = sinh(0.0)\nprint *, x\nend program t\n");
    assert_eq!(out, ["0"]);
}

#[test]
fn cosh_zero() {
    let out = run_prints("program t\nreal :: x\nx = cosh(0.0)\nprint *, x\nend program t\n");
    assert_eq!(out, ["1"]);
}

#[test]
fn tanh_zero() {
    let out = run_prints("program t\nreal :: x\nx = tanh(0.0)\nprint *, x\nend program t\n");
    assert_eq!(out, ["0"]);
}

// ── Power and log ────────────────────────────────────────────

#[test]
fn log10_100() {
    let out = run_prints("program t\nreal :: x\nx = log10(100.0)\nprint *, x\nend program t\n");
    assert_eq!(out, ["2"]);
}

#[test]
fn exp_zero() {
    compile_ok("program t\nreal :: x\nx = exp(0.0)\nprint *, x\nend program t\n");
}

// ── Integer intrinsics ────────────────────────────────────────

#[test]
fn min_three() {
    let out = run_prints("program t\nprint *, min(5, 3, 8)\nend program t\n");
    assert_eq!(out, ["3"]);
}

#[test]
fn max_three() {
    let out = run_prints("program t\nprint *, max(5, 3, 8)\nend program t\n");
    assert_eq!(out, ["8"]);
}

#[test]
fn min_four() {
    compile_ok("program t\nprint *, min(10, 20, 5, 15)\nend program t\n");
}

#[test]
fn max_four() {
    compile_ok("program t\nprint *, max(10, 20, 5, 15)\nend program t\n");
}

#[test]
fn abs_real() {
    let out = run_prints("program t\nreal :: x\nx = abs(-3.5)\nprint *, x\nend program t\n");
    assert_eq!(out, ["3.5"]);
}

// ── Bit operations ────────────────────────────────────────────

#[test]
fn iand_basic() {
    let out = run_prints("program t\nprint *, iand(255, 15)\nend program t\n");
    assert_eq!(out, ["15"]);
}

#[test]
fn ior_basic() {
    let out = run_prints("program t\nprint *, ior(240, 15)\nend program t\n");
    assert_eq!(out, ["255"]);
}

#[test]
fn ieor_basic() {
    let out = run_prints("program t\nprint *, ieor(255, 15)\nend program t\n");
    assert_eq!(out, ["240"]);
}

#[test]
fn ishft_left() {
    let out = run_prints("program t\nprint *, ishft(1, 4)\nend program t\n");
    assert_eq!(out, ["16"]);
}

#[test]
fn ishft_right() {
    let out = run_prints("program t\nprint *, ishft(256, -4)\nend program t\n");
    assert_eq!(out, ["16"]);
}

#[test]
fn ibset_bit() {
    let out = run_prints("program t\nprint *, ibset(0, 3)\nend program t\n");
    assert_eq!(out, ["8"]);
}

#[test]
fn ibclr_bit() {
    let out = run_prints("program t\nprint *, ibclr(15, 0)\nend program t\n");
    assert_eq!(out, ["14"]);
}

#[test]
fn not_basic() {
    let out = run_prints("program t\ninteger :: x\nx = not(0)\nprint *, x\nend program t\n");
    assert_eq!(out, ["-1"]);
}

// ── Type / kind intrinsics ────────────────────────────────────

#[test]
fn kind_int() {
    compile_ok("program t\nprint *, kind(0)\nend program t\n");
}

#[test]
fn kind_real() {
    compile_ok("program t\nprint *, kind(0.0)\nend program t\n");
}

#[test]
fn selected_int_kind() {
    compile_ok("program t\ninteger :: k\nk = selected_int_kind(9)\nprint *, k\nend program t\n");
}

#[test]
fn selected_real_kind() {
    compile_ok("program t\ninteger :: k\nk = selected_real_kind(15)\nprint *, k\nend program t\n");
}

// ── Numeric queries ───────────────────────────────────────────

#[test]
fn huge_int() {
    compile_ok("program t\nprint *, huge(0)\nend program t\n");
}

#[test]
fn tiny_real() {
    compile_ok("program t\nprint *, tiny(0.0)\nend program t\n");
}

#[test]
fn epsilon_real() {
    compile_ok("program t\nprint *, epsilon(0.0)\nend program t\n");
}

// ── Random ───────────────────────────────────────────────────

#[test]
fn random_number() {
    compile_ok(r#"
program test
    real :: r
    call random_number(r)
    print *, r >= 0.0 .and. r < 1.0
end program test
"#);
}

#[test]
fn random_seed() {
    compile_ok(r#"
program test
    integer :: seed(1) = [42]
    call random_seed(put=seed)
    print *, "ok"
end program test
"#);
}

// ── Date and time ─────────────────────────────────────────────

#[test]
fn date_and_time() {
    compile_ok(r#"
program test
    character(len=8) :: d
    character(len=10) :: t
    call date_and_time(date=d, time=t)
    print *, "got date"
end program test
"#);
}

// ── Merge ────────────────────────────────────────────────────

#[test]
fn merge_basic() {
    let out = run_prints("program t\nprint *, merge(1, 0, .true.)\nend program t\n");
    assert_eq!(out, ["1"]);
}

#[test]
fn merge_false() {
    let out = run_prints("program t\nprint *, merge(1, 0, .false.)\nend program t\n");
    assert_eq!(out, ["0"]);
}
