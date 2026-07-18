use super::helpers::run_prints;

#[test]
fn test_named_constant_initialization_uses_expression() {
    let out = run_prints(
        r#"
program test_named_constant_initialization
    real, parameter :: pi = 3.14159
    print *, nint(pi)
end program test_named_constant_initialization
"#,
    );

    assert_eq!(out, vec!["3"]);
}
