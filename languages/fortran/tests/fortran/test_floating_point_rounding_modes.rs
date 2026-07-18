use super::helpers::run_prints;

#[test]
fn test_floating_point_rounding_modes_nint() {
    let out = run_prints(
        r#"
program test_floating_point_rounding_modes
    real :: value
    value = 2.6
    print *, nint(value)
    value = -2.6
    print *, nint(value)
end program test_floating_point_rounding_modes
"#,
    );

    assert_eq!(out, vec!["3", "-3"]);
}
