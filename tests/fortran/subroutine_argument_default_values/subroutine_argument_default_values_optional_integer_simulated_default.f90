! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_optional_integer_simulated_default
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

program t
    if ((apply_default(5)) /= 6) then
    print *, "FAIL: want [6] got [", apply_default(5), "]"
    stop 1
end if
    if ((apply_default(5, 2)) /= 7) then
    print *, "FAIL: want [7] got [", apply_default(5, 2), "]"
    stop 1
end if
contains
    integer function apply_default(value, step)
        integer, intent(in) :: value
        integer, intent(in), optional :: step
        integer :: step_value
        if (present(step)) then
            step_value = step
        else
            step_value = 1
        end if
        apply_default = value + step_value
    end function apply_default
end program t
