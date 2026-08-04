! vybe-test: fortran/basics/implicit_none_multiple_assignment_steps
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    implicit none
    integer :: a
    integer :: b
    a = 1
    b = 2
    a = a + b
    if ((a) /= 3) then
    print *, "FAIL: want [3] got [", a, "]"
    stop 1
end if
    if ((a * b) /= 6) then
    print *, "FAIL: want [6] got [", a * b, "]"
    stop 1
end if
end program test
