! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_optional_out_with_default_mode
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

program subroutine_argument_default_values_optional_out_with_default_mode
    integer :: out_val
    call maybe_set(out_val)
    if ((out_val) /= 99) then
    print *, "FAIL: want [99] got [", out_val, "]"
    stop 1
end if
contains
    subroutine maybe_set(result, value)
        integer, intent(out) :: result
        integer, intent(in), optional :: value
        if (present(value)) then
            result = value
        else
            result = 99
        end if
    end subroutine maybe_set
end program subroutine_argument_default_values_optional_out_with_default_mode
