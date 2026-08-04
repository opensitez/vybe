! vybe-test: fortran/expressions/expr_sub_02
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a=5,b=2,c
c = a - b
if ((c) /= 3) then
    print *, "FAIL: want [3] got [", c, "]"
    stop 1
end if
end program p
