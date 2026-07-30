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

#[test]
fn test_generic_ambiguous_interface_errors_character_and_logical_dispatch() {
    let out = run_prints(
        r#"
program test_generic_ambiguous_interface_errors
    print *, magnitude('abc')
    print *, magnitude(.true.)

contains
    interface magnitude
        module procedure abs_char
        module procedure abs_log
    end interface

    integer function abs_char(v)
        character(len=*), intent(in) :: v
        abs_char = len_trim(v)
    end function

    integer function abs_log(v)
        logical, intent(in) :: v
        if (v) then
            abs_log = 1
        else
            abs_log = 0
        end if
    end function
end program test_generic_ambiguous_interface_errors
"#,
    );

    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn test_generic_ambiguous_interface_errors_real_integer_no_ambiguity() {
    let out = run_prints(
        r#"
program test_generic_ambiguous_interface_errors
    print *, magnitude(2)
    print *, nint(magnitude(-3.0))

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

    assert_eq!(out, vec!["2", "3"]);
}
