use super::helpers::run_prints;

#[test]
fn test_generic_interface_resolution_selects_integer_binding() {
    let out = run_prints(
        r#"
program test_generic_interface_resolution
    integer :: result
    result = add(2, 3)
    print *, result

contains
    interface add
        module procedure add_int
    end interface

    integer function add_int(a, b)
        integer, intent(in) :: a, b
        add_int = a + b
    end function
end program test_generic_interface_resolution
"#,
    );

    assert_eq!(out, vec!["5"]);
}
