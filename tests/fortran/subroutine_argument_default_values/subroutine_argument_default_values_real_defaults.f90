! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_real_defaults
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

program subroutine_argument_default_values_real_defaults
    if ((ratio(10.0)) /= 5) then
    print *, "FAIL: want [5] got [", ratio(10.0), "]"
    stop 1
end if
    if ((ratio(10.0, 4.0)) /= 2) then
    print *, "FAIL: want [2] got [", ratio(10.0, 4.0), "]"
    stop 1
end if
contains
    real function ratio(value, divisor)
        real, intent(in) :: value
        real, intent(in), optional :: divisor
        real :: d
        if (present(divisor)) then
            d = divisor
        else
            d = 2.0
        end if
        ratio = value / d
    end function ratio
end program subroutine_argument_default_values_real_defaults
