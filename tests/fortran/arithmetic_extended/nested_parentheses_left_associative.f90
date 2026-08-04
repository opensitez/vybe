! vybe-test: fortran/arithmetic_extended/nested_parentheses_left_associative
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((((1 + 2) * 3) + 4) /= 13) then
    print *, "FAIL: want [13] got [", ((1 + 2) * 3) + 4, "]"
    stop 1
end if
end program t
