use super::helpers::run_prints;

#[test]
fn test_do_construct_stop_conditions_exit_at_threshold() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    do i = 1, 10
        if (i > 4) exit
        total = total + i
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["10"]);
}
