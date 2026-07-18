use super::helpers::run_prints;

#[test]
fn test_deferred_length_operators_on_allocatable_character() {
    let out = run_prints(
        r#"
program test_deferred_length_operators
    character(len=:), allocatable :: text
    allocate(character(len=9) :: text)
    text = 'fortify-1'
    print *, len(text)
    print *, trim(text)
end program test_deferred_length_operators
"#,
    );

    assert_eq!(out, vec!["9", "fortify-1"]);
}
