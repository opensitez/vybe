! vybe-test: fortran/recursive_optional_arguments/recursive_optional_arguments_optional_direction_signals
! origin: languages/fortran/tests/fortran/test_recursive_optional_arguments.rs

program recursive_optional_arguments_optional_direction_signals
    if ((sign_walk(4)) /= 4) then
    print *, "FAIL: want [4] got [", sign_walk(4), "]"
    stop 1
end if
    if ((sign_walk(4, -1)) /= 4) then
    print *, "FAIL: want [4] got [", sign_walk(4, -1), "]"
    stop 1
end if
contains
    recursive integer function sign_walk(n, dir) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: dir
        integer :: d
        d = 1
        if (present(dir)) d = dir
        if (abs(d) /= 1) d = 1
        if (n <= 0) then
            out = 0
        else if (d == -1) then
            out = 1 + sign_walk(n - 1)
        else
            out = n
        end if
    end function sign_walk
end program recursive_optional_arguments_optional_direction_signals
