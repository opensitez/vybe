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

    assert_eq!(out, vec!["True"]);
}
