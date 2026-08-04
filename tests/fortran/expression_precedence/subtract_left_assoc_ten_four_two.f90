! vybe-test: fortran/expression_precedence/subtract_left_assoc_ten_four_two
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((10 - 4 - 2) /= 4) then
    print *, "FAIL: want [4] got [", 10 - 4 - 2, "]"
    stop 1
end if
end program t
