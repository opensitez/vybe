! vybe-test: fortran/expressions/expr_paren_07
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: x
x = (2 + 3) * 4
if ((x) /= 20) then
    print *, "FAIL: want [20] got [", x, "]"
    stop 1
end if
end program p
