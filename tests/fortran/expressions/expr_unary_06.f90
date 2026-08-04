! vybe-test: fortran/expressions/expr_unary_06
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a
a = -5
if ((a) /= -5) then
    print *, "FAIL: want [-5] got [", a, "]"
    stop 1
end if
end program p
