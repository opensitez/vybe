! vybe-test: fortran/arithmetic_extended/left_associative_division_chain
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((100 / 10 / 2) /= 5) then
    print *, "FAIL: want [5] got [", 100 / 10 / 2, "]"
    stop 1
end if
end program t
