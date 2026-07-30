use super::helpers::run_prints;
use super::helpers::compile_ok;

#[test]
fn test_kind_parameter_defaulting_selects_default_real_kind() {
    let out = run_prints(
        r#"
program test_kind_parameter_defaulting
    integer :: k
    k = selected_real_kind(6)
    print *, k
end program test_kind_parameter_defaulting
"#,
    );

    assert_eq!(out, ["8"]);
}

#[test]
fn test_selected_real_kind_defaulting_with_unavailable_range_is_valid() {
    let out = run_prints(
        r#"
program test_kind_parameter_defaulting
    print *, selected_real_kind(6)
    print *, selected_real_kind(6, 38)
    print *, selected_real_kind(15)
end program test_kind_parameter_defaulting
"#,
    );

    assert_eq!(out, ["8", "8", "8"]);
}

#[test]
fn test_selected_int_kind_is_usable_for_declarations_with_defaults() {
    compile_ok(
        r#"
program test_kind_parameter_defaulting
    integer, parameter :: i4 = selected_int_kind(9)
    integer(kind=i4) :: i
    real, parameter :: r4 = selected_real_kind(5)
    real(kind=r4) :: x
    integer :: j
    real :: y
    i = 1
    x = 1.0
    j = kind(i)
    y = real(j)
    print *, kind(i)
    print *, kind(x)
end program test_kind_parameter_defaulting
"#,
    );
}
