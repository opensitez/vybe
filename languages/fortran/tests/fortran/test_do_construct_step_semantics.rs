use super::helpers::run_prints;

#[test]
fn test_do_construct_step_semantics_decrements_with_custom_stride() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics
    integer :: i
    integer :: total
    total = 0
    do i = 10, 2, -4
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics
"#,
    );

    assert_eq!(out, vec!["18"]);
}
