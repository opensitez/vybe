use super::helpers::run_prints;

#[test]
fn test_pointer_reassociation_sequences_switch_targets() {
    let out = run_prints(
        r#"
program test_pointer_reassociation_sequences
    integer, target :: a
    integer, target :: b
    integer, pointer :: p

    a = 5
    b = 9
    p => a
    print *, p
    p => b
    print *, p
end program test_pointer_reassociation_sequences
"#,
    );

    assert_eq!(out, vec!["5", "9"]);
}
