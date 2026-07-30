use super::helpers::run_prints;

#[test]
fn test_ieee_signaling_exceptions_queries_support() {
    let out = run_prints(
        r#"
program test_ieee_signaling_exceptions
    use, intrinsic :: ieee_arithmetic
    logical :: ok
    ok = ieee_support_datatype(0.0)
    print *, ok
end program test_ieee_signaling_exceptions
"#,
    );

    assert_eq!(out.len(), 1);
    assert!(out[0].to_lowercase().contains("true"));
}

#[test]
fn test_ieee_signaling_exceptions_nan_classification() {
    let out = run_prints(
        r#"
program test_ieee_signaling_exceptions_nan_classification
    use, intrinsic :: ieee_arithmetic
    real :: x
    logical :: is_nan
    x = ieee_value(x, ieee_signaling_nan)
    is_nan = ieee_is_nan(x)
    print *, is_nan
    print *, ieee_is_finite(x)
end program test_ieee_signaling_exceptions_nan_classification
"#,
    );
    assert_eq!(out.len(), 2);
    assert!(out[0].to_lowercase().contains("true"));
    assert!(out[1].to_lowercase().contains("false"));
}

#[test]
fn test_ieee_signaling_exceptions_flag_state_query() {
    let out = run_prints(
        r#"
program test_ieee_signaling_exceptions_flag_state_query
    use, intrinsic :: ieee_exceptions
    logical :: flag_set
    call ieee_get_flag(ieee_divide_by_zero, flag_set)
    print *, flag_set
end program test_ieee_signaling_exceptions_flag_state_query
"#,
    );
    assert_eq!(out.len(), 1);
    assert!(out[0].to_lowercase().contains("false"));
}
