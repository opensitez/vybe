! vybe-test: fortran/expressions/expr_power_precedence_33
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a
a = -2 ** 3
if ((a) /= -8) then
    print *, "FAIL: want [-8] got [", a, "]"
    stop 1
end if
end program p
