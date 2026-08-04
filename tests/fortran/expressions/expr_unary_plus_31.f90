! vybe-test: fortran/expressions/expr_unary_plus_31
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a
a = +7
if ((a) /= 7) then
    print *, "FAIL: want [7] got [", a, "]"
    stop 1
end if
end program p
