use super::helpers::run_prints;

#[test]
fn test_named_constant_statement_functions_evaluate_with_constants() {
    let out = run_prints(
        r#"
program test_named_constant_statement_functions
    integer :: n
    integer :: cube
    n = 4
    cube(n) = n ** 3
    print *, cube(3)
end program test_named_constant_statement_functions
"#,
    );

    assert_eq!(out, vec!["27"]);
}
