use super::helpers::run_prints;

#[test]
fn test_coindexed_object_access_declares_single_image_access() {
    let out = run_prints(
        r#"
program test_coindexed_object_access
    integer, allocatable :: shared(:)
    integer :: this
    allocate(shared(1))
    shared(1) = 7
    this = this_image()
    print *, shared(1)
    print *, this
end program test_coindexed_object_access
"#,
    );

    assert_eq!(out, vec!["7", "1"]);
}
