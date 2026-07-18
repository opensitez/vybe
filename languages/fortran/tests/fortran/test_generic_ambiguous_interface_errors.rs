use super::helpers::run_prints;

#[test]
fn test_generic_ambiguous_interface_errors_with_distinct_specifics() {
    let out = run_prints(
        r#"
program test_generic_ambiguous_interface_errors
    print *, magnitude(3)
    print *, magnitude(2.5)

contains
    interface magnitude
        module procedure abs_int
        module procedure abs_real
    end interface

    integer function abs_int(v)
        integer, intent(in) :: v
        abs_int = abs(v)
    end function

    real function abs_real(v)
        real, intent(in) :: v
        abs_real = abs(v)
    end function
end program test_generic_ambiguous_interface_errors
"#,
    );

    assert_eq!(out, vec!["3", "2.5"]);
}
