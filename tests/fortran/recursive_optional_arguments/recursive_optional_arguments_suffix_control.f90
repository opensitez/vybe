! vybe-test: fortran/recursive_optional_arguments/recursive_optional_arguments_suffix_control
! origin: languages/fortran/tests/fortran/test_recursive_optional_arguments.rs

program recursive_optional_arguments_suffix_control
    if ((fold(4)) /= 10) then
    print *, "FAIL: want [10] got [", fold(4), "]"
    stop 1
end if
    if ((fold(4, 2)) /= 10) then
    print *, "FAIL: want [10] got [", fold(4, 2), "]"
    stop 1
end if
contains
    recursive integer function fold(n, step) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: step
        integer :: step_value
        step_value = 1
        if (present(step)) step_value = step
        if (n <= 0) then
            out = 0
        else
            out = n + fold(n - step_value, step_value)
        end if
    end function fold
end program recursive_optional_arguments_suffix_control
