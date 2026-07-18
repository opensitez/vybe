use super::helpers::run_prints;

#[test]
fn test_namespace_collisions_in_scope_resolves_inner_identifier() {
    let out = run_prints(
        r#"
program test_namespace_collisions_in_scope
    integer :: value
    value = 1
    block
        integer :: value
        value = 3
        print *, value
    end block
    print *, value
end program test_namespace_collisions_in_scope
"#,
    );

    assert_eq!(out, vec!["3", "1"]);
}
