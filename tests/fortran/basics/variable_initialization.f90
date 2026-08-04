! vybe-test: fortran/basics/variable_initialization
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    integer :: x = 10
    integer :: y = 20
    if ((x + y) /= 30) then
    print *, "FAIL: want [30] got [", x + y, "]"
    stop 1
end if
end program test
