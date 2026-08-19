! vybe-test: fortran/recursive_optional_arguments/recursive_optional_arguments_optional_limit_guard
! origin: languages/fortran/tests/fortran/test_recursive_optional_arguments.rs

program recursive_optional_arguments_optional_limit_guard
    if ((limited(5)) /= 5) then
    print *, "FAIL: want [5] got [", limited(5), "]"
    stop 1
end if
    if ((limited(5, 2)) /= 5) then
    print *, "FAIL: want [5] got [", limited(5, 2), "]"
    stop 1
end if
contains
    recursive integer function limited(n, limit) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: limit
        integer :: max_depth
        max_depth = 1
        if (present(limit)) max_depth = limit
        if (n <= 0 .or. n > 10*max_depth) then
            out = 0
        else
            out = 1 + limited(n - 1, max_depth)
        end if
    end function limited
end program recursive_optional_arguments_optional_limit_guard
