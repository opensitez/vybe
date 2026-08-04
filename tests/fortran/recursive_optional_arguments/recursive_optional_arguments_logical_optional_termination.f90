! vybe-test: fortran/recursive_optional_arguments/recursive_optional_arguments_logical_optional_termination
! origin: languages/fortran/tests/fortran/test_recursive_optional_arguments.rs

program recursive_optional_arguments_logical_optional_termination
    if ((parity(6)) /= 21) then
    print *, "FAIL: want [21] got [", parity(6), "]"
    stop 1
end if
    if ((parity(6, .false.)) /= 9) then
    print *, "FAIL: want [9] got [", parity(6, .false.), "]"
    stop 1
end if
contains
    recursive integer function parity(n, even_only) result(out)
        integer, intent(in) :: n
        logical, optional, intent(in) :: even_only
        logical :: only_even
        only_even = .true.
        if (present(even_only)) only_even = even_only
        if (n <= 0) then
            out = 0
        else if (mod(n,2) == 0 .or. .not. only_even) then
            out = n + parity(n - 1, only_even)
        else
            out = parity(n - 1, only_even)
        end if
    end function parity
end program recursive_optional_arguments_logical_optional_termination
