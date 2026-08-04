! vybe-test: fortran/arithmetic/power_small
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((3 ** 3) /= 27) then
    print *, "FAIL: want [27] got [", 3 ** 3, "]"
    stop 1
end if
end program t
