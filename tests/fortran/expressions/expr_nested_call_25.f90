! vybe-test: fortran/expressions/expr_nested_call_25
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
if ((max(1, min(2,3))) /= 2) then
    print *, "FAIL: want [2] got [", max(1, min(2,3)), "]"
    stop 1
end if
end program p
