use super::helpers::run_prints;

#[test]
fn test_pointer_vectorized_assignment_copies_slice() {
    let out = run_prints(
        r#"
program test_pointer_vectorized_assignment
    integer, target :: src(3)
    integer, target :: dst(3)
    integer, pointer :: psrc(:)
    integer, pointer :: pdst(:)

    src = (/1, 2, 3/)
    psrc => src
    pdst => dst
    pdst = psrc
    print *, pdst(2)
end program test_pointer_vectorized_assignment
"#,
    );

    assert_eq!(out, vec!["2"]);
}
