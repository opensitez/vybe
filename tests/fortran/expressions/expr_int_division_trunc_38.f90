! vybe-test: fortran/expressions/expr_int_division_trunc_38
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
if ((7 / 2) /= 3) then
    print *, "FAIL: want [3] got [", 7 / 2, "]"
    stop 1
end if
if ((-17 / 5) /= -3) then
    print *, "FAIL: want [-3] got [", -17 / 5, "]"
    stop 1
end if
end program p
