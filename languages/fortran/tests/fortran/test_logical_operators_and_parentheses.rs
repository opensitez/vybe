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

#[test]
fn test_logical_eqv_neqv_precedence() {
    let out = run_prints(
        r#"
program test_logical_eqv_neqv_precedence
    print *, .true. .eqv. .true.
    print *, .true. .neqv. .true.
    print *, .true. .or. (.false. .eqv. .true.)
end program test_logical_eqv_neqv_precedence
"#,
    );

    assert_eq!(out, vec!["True", "False", "True"]);
}

#[test]
fn test_not_with_parentheses() {
    let out = run_prints(
        r#"
program test_not_with_parentheses
    print *, .not. (.true. .and. .false.)
    print *, .not. (.false. .or. .false.)
    print *, (.not. .true.) .and. .true.
end program test_not_with_parentheses
"#,
    );

    assert_eq!(out, vec!["True", "True", "False"]);
}
