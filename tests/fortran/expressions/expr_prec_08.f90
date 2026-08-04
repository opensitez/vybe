! vybe-test: fortran/expressions/expr_prec_08
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: x
x = 2 + 3 * 4
if ((x) /= 14) then
    print *, "FAIL: want [14] got [", x, "]"
    stop 1
end if
end program p
