use super::helpers::run_prints;

#[test]
fn test_include_directive_line_continuations_use_ampersand() {
    let out = run_prints(
        r#"
program test_include_directive_line_continuations
    integer :: value
    value = 1 + 2 + &
            3 + 4
    print *, value
end program test_include_directive_line_continuations
"#,
    );

    assert_eq!(out, vec!["10"]);
}
