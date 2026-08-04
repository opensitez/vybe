! vybe-test: fortran/arithmetic/power_is_right_associative_for_exponent_chain
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((2 ** 3 ** 2) /= 512) then
    print *, "FAIL: want [512] got [", 2 ** 3 ** 2, "]"
    stop 1
end if
end program t
