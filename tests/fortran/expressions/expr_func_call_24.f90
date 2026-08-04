! vybe-test: fortran/expressions/expr_func_call_24
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
if ((abs(-3)) /= 3) then
    print *, "FAIL: want [3] got [", abs(-3), "]"
    stop 1
end if
end program p
