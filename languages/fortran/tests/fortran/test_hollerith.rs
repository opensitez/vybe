use super::helpers::{compile_ok, run_prints};

// ── Hollerith in FORMAT statements ───────────────────────────
// Hollerith nH... is F66/F77 legacy. Most F90+ compilers accept
// it in FORMAT strings for backward compatibility.

#[test]
fn hollerith_in_format() {
    compile_ok(
        r#"
program test
    write(*, 100)
100 format(5Hhello)
end program test
"#,
    );
}

#[test]
fn hollerith_with_integer() {
    compile_ok(
        r#"
program test
    integer :: n = 42
    write(*, 100) n
100 format(5Hval= , I4)
end program test
"#,
    );
}

#[test]
fn hollerith_single_char() {
    compile_ok(
        r#"
program test
    write(*, 100)
100 format(1Hx)
end program test
"#,
    );
}

#[test]
fn hollerith_space() {
    compile_ok(
        r#"
program test
    integer :: a = 1, b = 2
    write(*, 100) a, b
100 format(I3, 1H , I3)
end program test
"#,
    );
}

#[test]
fn hollerith_multiword() {
    compile_ok(
        r#"
program test
    write(*, 100)
100 format(13Hhello, world!)
end program test
"#,
    );
}

#[test]
fn hollerith_newline_equivalent() {
    compile_ok(
        r#"
program test
    write(*, 100)
100 format(4Hline)
    write(*, 200)
200 format(4Htwo!)
end program test
"#,
    );
}

// ── Hollerith in DATA statements ─────────────────────────────

#[test]
fn hollerith_in_data() {
    compile_ok(
        r#"
program test
    integer :: word
    data word /4HABCD/
    print *, 'ok'
end program test
"#,
    );
}

#[test]
fn hollerith_data_two_words() {
    compile_ok(
        r#"
program test
    integer :: w1, w2
    data w1 /4HTEST/, w2 /4HDATA/
    print *, 'ok'
end program test
"#,
    );
}

// ── Hollerith as subroutine argument ─────────────────────────

#[test]
fn hollerith_as_argument() {
    compile_ok(
        r#"
program test
    call show(5Hhello)
contains
    subroutine show(msg)
        integer, intent(in) :: msg
        print *, 'received'
    end subroutine show
end program test
"#,
    );
}

// ── Hollerith assigned to integer variable ────────────────────

#[test]
fn hollerith_assigned_to_integer() {
    compile_ok(
        r#"
program test
    integer :: tag
    tag = 4HTEST
    print *, 'ok'
end program test
"#,
    );
}

#[test]
fn hollerith_assigned_to_real() {
    compile_ok(
        r#"
program test
    real :: tag
    tag = 4HTEST
    print *, 'ok'
end program test
"#,
    );
}

// ── Hollerith in COMMON ───────────────────────────────────────

#[test]
fn hollerith_in_common() {
    compile_ok(
        r#"
program test
    integer :: label
    common /info/ label
    data label /4HINFO/
    print *, 'ok'
end program test
"#,
    );
}

// ── Hollerith with padding ────────────────────────────────────

#[test]
fn hollerith_padded_shorter() {
    compile_ok(
        r#"
program test
    integer :: w
    w = 2Hhi
    print *, 'ok'
end program test
"#,
    );
}

// ── H-format in write ─────────────────────────────────────────

#[test]
fn h_format_label_and_value() {
    compile_ok(
        r#"
program test
    real :: x = 3.14
    write(*, 10) x
10  format(7Hresult=, F6.2)
end program test
"#,
    );
}

#[test]
fn h_format_inside_repeat() {
    compile_ok(
        r#"
program test
    write(*, 10) 1, 2, 3
10  format(3(2Hv=, I2, 1H ))
end program test
"#,
    );
}

#[test]
fn hollerith_in_format_runtime() {
    let out = run_prints(
        r#"
program test
    write(*, 100)
100 format(5Hhello)
end program test
"#,
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn hollerith_with_integer_runtime() {
    let out = run_prints(
        r#"
program test
    integer :: n = 42
    write(*, 100) n
100 format(5Hval= , I4)
end program test
"#,
    );
    assert!(out[0].contains("val="));
    assert!(out[0].contains("42"));
}

#[test]
fn hollerith_space_runtime() {
    let out = run_prints(
        r#"
program test
    integer :: a = 1, b = 2
    write(*, 100) a, b
100 format(I3, 1H , I3)
end program test
"#,
    );
    assert!(out[0].contains("1"));
    assert!(out[0].contains("2"));
}
