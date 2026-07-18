use super::helpers::run_prints;

#[test]
fn test_derived_type_copy_out_copy_in_assigns_components() {
    let out = run_prints(
        r#"
program test_derived_type_copy_out_copy_in
    type :: pair
        integer :: a
        integer :: b
    end type

    type(pair) :: source
    type(pair) :: dest

    source%a = 2
    source%b = 5
    dest = source

    print *, dest%a
    print *, dest%b
end program test_derived_type_copy_out_copy_in
"#,
    );

    assert_eq!(out, vec!["2", "5"]);
}
