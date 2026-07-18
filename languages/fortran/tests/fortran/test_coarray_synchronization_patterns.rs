use super::helpers::run_prints;

#[test]
fn test_coarray_synchronization_patterns_sync_all_anchor() {
    let out = run_prints(
        r#"
program test_coarray_synchronization_patterns
    integer :: value
    value = 11
    sync all
    print *, value
end program test_coarray_synchronization_patterns
"#,
    );

    assert_eq!(out, vec!["11"]);
}
