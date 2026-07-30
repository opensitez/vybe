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

#[test]
fn test_generic_interface_resolution_selects_real_and_integer_bindings() {
    let out = run_prints(
        r#"
program test_generic_interface_resolution_selects_real_and_integer_bindings
    print *, add(2, 3)
    print *, nint(add(2.0, 3.5))

contains
    interface add
        module procedure add_int
        module procedure add_real
    end interface

    integer function add_int(a, b)
        integer, intent(in) :: a, b
        add_int = a + b
    end function

    real function add_real(a, b)
        real, intent(in) :: a, b
        add_real = a + b
    end function
end program test_generic_interface_resolution_selects_real_and_integer_bindings
"#,
    );

    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn test_generic_interface_resolution_subroutine_call_dispatch() {
    let out = run_prints(
        r#"
program test_generic_interface_resolution_subroutine_call_dispatch
    call scale_out(2)
    call scale_out(3.0)

contains
    interface scale_out
        module procedure scale_i
        module procedure scale_r
    end interface

    subroutine scale_i(value)
        integer, intent(in) :: value
        print *, value * 10
    end subroutine

    subroutine scale_r(value)
        real, intent(in) :: value
        print *, nint(value * 10.0)
    end subroutine
end program test_generic_interface_resolution_subroutine_call_dispatch
"#,
    );

    assert_eq!(out, vec!["20", "30"]);
}

#[test]
fn test_generic_interface_resolution_character_and_logical_specifics() {
    let out = run_prints(
        r#"
program test_generic_interface_resolution_character_and_logical_specifics
    print *, size_or_truth("abc")
    print *, size_or_truth(.true.)

contains
    interface size_or_truth
        module procedure size_text
        module procedure truth_int
    end interface

    integer function size_text(value)
        character(len=*), intent(in) :: value
        size_text = len_trim(value)
    end function

    integer function truth_int(value)
        logical, intent(in) :: value
        if (value) then
            truth_int = 1
        else
            truth_int = 0
        end if
    end function
end program test_generic_interface_resolution_character_and_logical_specifics
"#,
    );

    assert_eq!(out, vec!["3", "1"]);
}
