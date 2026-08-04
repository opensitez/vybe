! vybe-test: fortran/basics/integer_variable
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    integer :: x
    x = 42
    if ((x) /= 42) then
    print *, "FAIL: want [42] got [", x, "]"
    stop 1
end if
end program test
