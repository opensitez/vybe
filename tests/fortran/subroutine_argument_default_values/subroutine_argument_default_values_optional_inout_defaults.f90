! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_optional_inout_defaults
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

program subroutine_argument_default_values_optional_inout_defaults
    integer :: x
    x = 8
    call scale_default(x)
    if ((x) /= 16) then
    print *, "FAIL: want [16] got [", x, "]"
    stop 1
end if
    call scale_default(x, addend=5)
    if ((x) /= 21) then
    print *, "FAIL: want [21] got [", x, "]"
    stop 1
end if
contains
    subroutine scale_default(value, addend)
        integer, intent(inout) :: value
        integer, intent(in), optional :: addend
        if (present(addend)) then
            value = value + addend
        else
            value = value * 2
        end if
    end subroutine scale_default
end program subroutine_argument_default_values_optional_inout_defaults
