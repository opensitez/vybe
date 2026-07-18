use super::helpers::run_prints;

#[test]
fn test_parameter_expression_constant_folding_precomputes_arithmetic() {
    let out = run_prints(
        r#"
program test_parameter_expression_constant_folding
    integer, parameter :: a = 2 + 3 * 4 - 1
    integer, parameter :: b = merge(10, 1, a > 10)
    print *, a
    print *, b
end program test_parameter_expression_constant_folding
"#,
    );

    assert_eq!(out, vec!["13", "1"]);
}
