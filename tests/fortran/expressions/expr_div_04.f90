! vybe-test: fortran/expressions/expr_div_04
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
real :: a=8.0,b=2.0,c
c = a / b
if ((c) /= 4) then
    print *, "FAIL: want [4] got [", c, "]"
    stop 1
end if
end program p
