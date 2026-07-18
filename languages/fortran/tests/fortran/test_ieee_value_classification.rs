use super::helpers::run_prints;

#[test]
fn test_ieee_value_classification_checks_finite() {
    let out = run_prints(
        r#"
program test_ieee_value_classification
    use, intrinsic :: ieee_arithmetic
    real :: value
    value = 0.0
    print *, ieee_is_finite(value)
    print *, ieee_is_nan(value / 0.0)
end program test_ieee_value_classification
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}
