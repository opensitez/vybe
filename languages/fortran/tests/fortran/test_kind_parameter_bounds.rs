use super::helpers::{compile_ok, run_prints};

#[test]
fn test_kind_parameter_bounds_compute_integer_range() {
    let out = run_prints(
        r#"
program test_kind_parameter_bounds
    integer :: small
    integer :: medium
    small = selected_int_kind(4)
    medium = selected_int_kind(8)
    print *, small
    print *, medium
end program test_kind_parameter_bounds
"#,
    );

    assert_eq!(out, ["8", "8"]);
}

#[test]
fn test_selected_int_kind_falls_back_to_unavailable() {
    let out = run_prints(
        r#"
program test_kind_parameter_bounds
    print *, selected_int_kind(1000)
end program test_kind_parameter_bounds
"#,
    );

    assert_eq!(out, ["-1"]);
}

#[test]
fn test_selected_real_kind_boundaries() {
    let out = run_prints(
        r#"
program test_kind_parameter_bounds
    integer :: p
    p = selected_real_kind(6, 37)
    print *, p
end program test_kind_parameter_bounds
"#,
    );

    assert_eq!(out, ["8"]);
}

#[test]
fn test_selected_real_kind_with_unavailable_bounds_is_valid() {
    compile_ok(
        r#"
program test_kind_parameter_bounds
    integer :: p
    p = selected_real_kind(999, 999)
    p = selected_real_kind(6, 0)
    print *, p
end program test_kind_parameter_bounds
"#,
    );
}

#[test]
fn test_kind_parameters_are_reusable_in_declarations() {
    compile_ok(
        r#"
program test_kind_parameter_bounds
    integer, parameter :: ik = selected_int_kind(9)
    real, parameter :: rk = selected_real_kind(15)
    integer(kind=ik) :: i
    real(kind=rk) :: x
    i = 7
    x = 2.5
    print *, kind(i)
    print *, kind(x)
end program test_kind_parameter_bounds
"#,
    );
}
