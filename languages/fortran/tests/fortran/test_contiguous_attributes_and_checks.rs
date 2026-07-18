use super::helpers::run_prints;

#[test]
fn test_contiguous_attributes_and_checks_valid_contiguous_section() {
    let out = run_prints(
        r#"
program test_contiguous_attributes_and_checks
    implicit none
    real :: values(5)
    values = (/1.0, 2.0, 3.0, 4.0, 5.0/)
    call inspect(values)

contains
    subroutine inspect(a)
        real, contiguous, intent(in) :: a(:)
        print *, size(a)
    end subroutine inspect
end program test_contiguous_attributes_and_checks
"#,
    );

    assert_eq!(out, vec!["5"]);
}
