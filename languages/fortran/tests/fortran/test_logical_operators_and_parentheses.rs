use super::helpers::run_prints;

#[test]
fn test_logical_operators_and_parentheses_mix_and_or() {
    let out = run_prints(
        r#"
program test_logical_operators_and_parentheses
    logical :: a
    logical :: b
    logical :: c
    a = .true.
    b = .false.
    c = .true.
    print *, (a .and. .not. b) .or. c
    print *, (a .and. (b .or. c))
end program test_logical_operators_and_parentheses
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}
