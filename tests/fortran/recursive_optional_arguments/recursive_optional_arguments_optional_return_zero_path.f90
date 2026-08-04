! vybe-test: fortran/recursive_optional_arguments/recursive_optional_arguments_optional_return_zero_path
! origin: languages/fortran/tests/fortran/test_recursive_optional_arguments.rs

program recursive_optional_arguments_optional_return_zero_path
    if ((capped(9)) /= 45) then
    print *, "FAIL: want [45] got [", capped(9), "]"
    stop 1
end if
    if ((capped(9, 4)) /= 10) then
    print *, "FAIL: want [10] got [", capped(9, 4), "]"
    stop 1
end if
contains
    recursive integer function capped(n, cap) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: cap
        integer :: m
        if (present(cap)) m = cap
        if (n <= 0) then
            out = 0
        else if (present(cap) .and. n > m) then
            out = m
        else
            out = n + capped(n - 1, cap)
        end if
    end function capped
end program recursive_optional_arguments_optional_return_zero_path
