use super::helpers::run_prints;

#[test]
fn test_coarray_sync_errors_checks_non_terminating_errors() {
    let out = run_prints(
        r#"
program test_coarray_sync_errors
    integer :: status
    status = 0
    if (status == 0) then
        print *, 0
    else
        print *, 1
    end if
end program test_coarray_sync_errors
"#,
    );

    assert_eq!(out, vec!["0"]);
}
