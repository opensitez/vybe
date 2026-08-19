! vybe-test: fortran/recursive_optional_arguments/recursive_optional_arguments_default_chain
! origin: languages/fortran/tests/fortran/test_recursive_optional_arguments.rs

program recursive_optional_arguments_default_chain
    if ((walk(3)) /= 6) then
    print *, "FAIL: want [6] got [", walk(3), "]"
    stop 1
end if
    if ((walk(3, 2)) /= 6) then
    print *, "FAIL: want [6] got [", walk(3, 2), "]"
    stop 1
end if
contains
    recursive integer function walk(n, step) result(out)
        integer, intent(in) :: n
        integer, optional, intent(in) :: step
        integer :: stride
        if (present(step)) then
            stride = step
        else
            stride = 1
        end if
        if (n <= 0) then
            out = 0
        else
            out = n + walk(n - 1, stride)
        end if
    end function walk
end program recursive_optional_arguments_default_chain
