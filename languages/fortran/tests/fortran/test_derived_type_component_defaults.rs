use super::helpers::run_prints;

#[test]
fn test_derived_type_component_defaults_apply_initial_values() {
    let out = run_prints(
        r#"
program test_derived_type_component_defaults
    type :: item
        integer :: x = 3
        logical :: enabled = .true.
    end type

    type(item) :: a
    if (a%enabled) then
        print *, a%x
    else
        print *, -1
    end if
end program test_derived_type_component_defaults
"#,
    );

    assert_eq!(out, vec!["3"]);
}
