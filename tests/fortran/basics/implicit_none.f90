! vybe-test: fortran/basics/implicit_none
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    implicit none
    integer :: x
    x = 42
    if ((x) /= 42) then
    print *, "FAIL: want [42] got [", x, "]"
    stop 1
end if
end program test
