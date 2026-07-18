use super::helpers::run_prints;

#[test]
fn test_optional_argument_association_present_keyword() {
    let out = run_prints(
        r#"
program test_optional_argument_association
    call show(5)
    call show(5, 2)

contains
    subroutine show(a, b)
        integer, intent(in) :: a
        integer, optional, intent(in) :: b
        if (present(b)) then
            print *, a + b
        else
            print *, a
        end if
    end subroutine
end program test_optional_argument_association
"#,
    );

    assert_eq!(out, vec!["5", "7"]);
}
