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

    assert_eq!(out.len(), 2);
    assert!(out[0].to_lowercase().contains("true"));
    assert!(out[1].to_lowercase().contains("false"));
}

#[test]
fn test_ieee_value_classification_infinity_vs_nan() {
    let out = run_prints(
        r#"
program test_ieee_value_classification_infinity_vs_nan
    use, intrinsic :: ieee_arithmetic
    real :: infv
    real :: nanv
    infv = ieee_value(infv, ieee_positive_inf)
    nanv = ieee_value(nanv, ieee_quiet_nan)
    print *, ieee_is_finite(infv)
    print *, ieee_is_nan(nanv)
    print *, ieee_class(infv) == ieee_positive_inf
    print *, ieee_class(nanv) == ieee_quiet_nan
end program test_ieee_value_classification_infinity_vs_nan
"#,
    );
    assert_eq!(out.len(), 4);
    assert!(out[0].to_lowercase().contains("false"));
    assert!(out[1].to_lowercase().contains("true"));
    assert!(out[2].to_lowercase().contains("true"));
    assert!(out[3].to_lowercase().contains("true"));
}
