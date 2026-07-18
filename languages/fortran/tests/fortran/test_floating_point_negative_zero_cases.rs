use super::helpers::run_prints;

#[test]
fn test_floating_point_negative_zero_cases_preserves_sign() {
    let out = run_prints(
        r#"
program test_floating_point_negative_zero_cases
    real :: x
    x = -0.0
    print *, sign(1.0, x)
    print *, x == 0.0
end program test_floating_point_negative_zero_cases
"#,
    );

    assert_eq!(out, vec!["-1", "True"]);
}
