use super::helpers::run_prints;

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

    assert_eq!(out.len(), 2);
}
