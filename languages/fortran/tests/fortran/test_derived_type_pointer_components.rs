use super::helpers::run_prints;

#[test]
fn test_derived_type_pointer_components_reference_and_read() {
    let out = run_prints(
        r#"
program test_derived_type_pointer_components
    type :: container
        integer, pointer :: values(:)
    end type

    integer, target :: storage(3)
    type(container) :: box

    storage = (/10, 20, 30/)
    box%values => storage
    print *, box%values(2)
end program test_derived_type_pointer_components
"#,
    );

    assert_eq!(out, vec!["20"]);
}
