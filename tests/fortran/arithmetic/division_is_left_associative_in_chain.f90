! vybe-test: fortran/arithmetic/division_is_left_associative_in_chain
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((100 / 10 / 2) /= 5) then
    print *, "FAIL: want [5] got [", 100 / 10 / 2, "]"
    stop 1
end if
end program t
