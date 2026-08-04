! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_type_defaults_via_optional
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

program subroutine_argument_default_values_type_defaults_via_optional
    type item
        integer :: x
    end type item
    type(item) :: base
    if ((set_item(base, 4)%x) /= 4) then
    print *, "FAIL: want [4] got [", set_item(base, 4)%x, "]"
    stop 1
end if
contains
    function set_item(base, x) result(out)
        type(item), intent(in) :: base
        integer, intent(in), optional :: x
        type(item) :: out
        if (present(x)) then
            out%x = base%x + x
        else
            out%x = base%x
        end if
    end function set_item
end program subroutine_argument_default_values_type_defaults_via_optional
