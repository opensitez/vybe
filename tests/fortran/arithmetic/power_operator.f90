! vybe-test: fortran/arithmetic/power_operator
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((2 ** 10) /= 1024) then
    print *, "FAIL: want [1024] got [", 2 ** 10, "]"
    stop 1
end if
end program t
