! vybe-test: fortran/arithmetic_extended/chained_unary_plus_minus_chain
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((10 - -5) /= 15) then
    print *, "FAIL: want [15] got [", 10 - -5, "]"
    stop 1
end if
end program t
