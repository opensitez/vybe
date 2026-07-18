use super::helpers::run_prints;

#[test]
fn test_coarray_component_initialization_defaults() {
    let out = run_prints(
        r#"
program test_coarray_component_initialization
    type :: endpoint
        integer :: value
        integer, allocatable :: values(:)
    end type endpoint

    type(endpoint) :: x
    allocate(x%values(3))
    x%values = (/1, 2, 3/)
    x%value = x%values(2)

    print *, x%value
    print *, x%values(3)
end program test_coarray_component_initialization
"#,
    );

    assert_eq!(out, vec!["2", "3"]);
}
