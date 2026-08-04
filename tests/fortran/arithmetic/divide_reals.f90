! vybe-test: fortran/arithmetic/divide_reals
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if (abs((10.0 / 4.0) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", 10.0 / 4.0, "]"
    stop 1
end if
end program t
