! vybe-test: fortran/basics/multiple_declarations
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    integer :: a = 1, b = 2, c = 3
    if ((a + b + c) /= 6) then
    print *, "FAIL: want [6] got [", a + b + c, "]"
    stop 1
end if
end program test
