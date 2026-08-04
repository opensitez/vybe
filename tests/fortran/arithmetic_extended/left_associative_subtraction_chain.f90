! vybe-test: fortran/arithmetic_extended/left_associative_subtraction_chain
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((8 - 3 - 2) /= 3) then
    print *, "FAIL: want [3] got [", 8 - 3 - 2, "]"
    stop 1
end if
end program t
