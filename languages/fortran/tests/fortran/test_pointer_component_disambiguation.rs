use super::helpers::run_prints;

#[test]
fn test_pointer_component_disambiguation_resolves_target_member() {
    let out = run_prints(
        r#"
program test_pointer_component_disambiguation
    type :: node
        integer :: a
    end type

    type(node), target :: n
    type(node), pointer :: p
    n%a = 8
    p => n
    print *, p%a
end program test_pointer_component_disambiguation
"#,
    );

    assert_eq!(out, vec!["8"]);
}
