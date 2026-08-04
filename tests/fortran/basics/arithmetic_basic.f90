! vybe-test: fortran/basics/arithmetic_basic
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    integer :: a, b
    a = 10
    b = 3
    if ((a + b) /= 13) then
    print *, "FAIL: want [13] got [", a + b, "]"
    stop 1
end if
    if ((a - b) /= 7) then
    print *, "FAIL: want [7] got [", a - b, "]"
    stop 1
end if
    if ((a * b) /= 30) then
    print *, "FAIL: want [30] got [", a * b, "]"
    stop 1
end if
end program test
