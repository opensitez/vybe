! vybe-test: fortran/expressions/expr_mul_03
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a=3,b=4,c
c = a * b
if ((c) /= 12) then
    print *, "FAIL: want [12] got [", c, "]"
    stop 1
end if
end program p
