use super::helpers::run_prints;

#[test]
fn test_do_while_progress_guarantees_monotonic_counter() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees
    integer :: i
    integer :: total
    i = 0
    total = 0
    do while (i < 5)
        i = i + 1
        total = total + i
    end do
    print *, i
    print *, total
end program test_do_while_progress_guarantees
"#,
    );

    assert_eq!(out, vec!["5", "15"]);
}
