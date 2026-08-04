! vybe-test: fortran/recursive_optional_arguments/recursive_optional_arguments_optional_weighted_accumulate
! origin: languages/fortran/tests/fortran/test_recursive_optional_arguments.rs

program recursive_optional_arguments_optional_weighted_accumulate
    if ((weighted(3)) /= 6) then
    print *, "FAIL: want [6] got [", weighted(3), "]"
    stop 1
end if
    if ((weighted(3, 10)) /= 33) then
    print *, "FAIL: want [33] got [", weighted(3, 10), "]"
    stop 1
end if
contains
    recursive integer function weighted(n, weight) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: weight
        integer :: w
        w = 1
        if (present(weight)) w = weight
        if (n <= 0) then
            out = 0
        else
            out = n * w + weighted(n - 1)
        end if
    end function weighted
end program recursive_optional_arguments_optional_weighted_accumulate
