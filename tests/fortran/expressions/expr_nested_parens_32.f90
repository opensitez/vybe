! vybe-test: fortran/expressions/expr_nested_parens_32
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a
a = (10 - 4) * (3 + 1)
if ((a) /= 24) then
    print *, "FAIL: want [24] got [", a, "]"
    stop 1
end if
end program p
