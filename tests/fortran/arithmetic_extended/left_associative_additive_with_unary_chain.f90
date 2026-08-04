! vybe-test: fortran/arithmetic_extended/left_associative_additive_with_unary_chain
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((20 - 5 + -3) /= 12) then
    print *, "FAIL: want [12] got [", 20 - 5 + -3, "]"
    stop 1
end if
end program t
