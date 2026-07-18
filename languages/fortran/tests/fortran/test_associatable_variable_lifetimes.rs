use super::helpers::run_prints;

#[test]
fn test_associatable_variable_lifetimes_preserves_outer_binding() {
    let out = run_prints(
        r#"
program test_associatable_variable_lifetimes
    implicit none
    integer :: base
    integer :: result
    base = 4
    associate(value => base)
        value = value + 1
        result = value
    end associate
    print *, result
    print *, base
end program test_associatable_variable_lifetimes
"#,
    );

    assert_eq!(out, vec!["5", "5"]);
}
