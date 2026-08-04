! vybe-test: fortran/expression_precedence/divide_before_add_twelve_over_three
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((12 / 3 + 1) /= 5) then
    print *, "FAIL: want [5] got [", 12 / 3 + 1, "]"
    stop 1
end if
end program t
