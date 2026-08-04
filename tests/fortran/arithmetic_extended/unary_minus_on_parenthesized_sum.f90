! vybe-test: fortran/arithmetic_extended/unary_minus_on_parenthesized_sum
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((-(3 + 4)) /= -7) then
    print *, "FAIL: want [-7] got [", -(3 + 4), "]"
    stop 1
end if
end program t
