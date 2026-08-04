! vybe-test: fortran/basics/exponentiation
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    if ((2 ** 10) /= 1024) then
    print *, "FAIL: want [1024] got [", 2 ** 10, "]"
    stop 1
end if
end program test
