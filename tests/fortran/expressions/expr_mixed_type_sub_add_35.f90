! vybe-test: fortran/expressions/expr_mixed_type_sub_add_35
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
if ((1 + 2.0 + 3) /= 6) then
    print *, "FAIL: want [6] got [", 1 + 2.0 + 3, "]"
    stop 1
end if
end program p
