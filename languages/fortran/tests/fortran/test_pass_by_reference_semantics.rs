use super::helpers::run_prints;

#[test]
fn test_pass_by_reference_semantics_mutates_caller_variable() {
    let out = run_prints(
        r#"
program test_pass_by_reference_semantics
    integer :: value
    value = 1
    call bump(value)
    print *, value

contains
    subroutine bump(x)
        integer, intent(inout) :: x
        x = x + 5
    end subroutine
end program test_pass_by_reference_semantics
"#,
    );

    assert_eq!(out, vec!["6"]);
}
