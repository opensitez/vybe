! vybe-test: fortran/arithmetic_extended/subtraction_before_multiplication
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((10 - 3 * 2) /= 4) then
    print *, "FAIL: want [4] got [", 10 - 3 * 2, "]"
    stop 1
end if
end program t
