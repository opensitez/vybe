! vybe-test: fortran/arithmetic/subtract_integers
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((10 - 3) /= 7) then
    print *, "FAIL: want [7] got [", 10 - 3, "]"
    stop 1
end if
end program t
