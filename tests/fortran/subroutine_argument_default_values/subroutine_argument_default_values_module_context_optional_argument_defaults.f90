! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_module_context_optional_argument_defaults
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

module opt_default_module
contains
    integer function weighted(value, bonus)
        integer, intent(in) :: value
        integer, intent(in), optional :: bonus
        if (present(bonus)) then
            weighted = value + bonus
        else
            weighted = value + 10
        end if
    end function weighted
end module opt_default_module

program subroutine_argument_default_values_module_context_optional_argument_defaults
    use opt_default_module
    if ((weighted(7)) /= 17) then
    print *, "FAIL: want [17] got [", weighted(7), "]"
    stop 1
end if
    if ((weighted(7, 2)) /= 9) then
    print *, "FAIL: want [9] got [", weighted(7, 2), "]"
    stop 1
end if
end program subroutine_argument_default_values_module_context_optional_argument_defaults
