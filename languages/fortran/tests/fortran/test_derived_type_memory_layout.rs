use super::helpers::run_prints;

#[test]
fn test_derived_type_memory_layout_reports_component_storage() {
    let out = run_prints(
        r#"
program test_derived_type_memory_layout
    type :: item
        integer :: a
        real :: b
    end type

    type(item) :: v
    print *, storage_size(v)
    print *, storage_size(v%a)
    print *, storage_size(v%b)
end program test_derived_type_memory_layout
"#,
    );

    assert_eq!(out, vec!["64", "32", "32"]);
}
