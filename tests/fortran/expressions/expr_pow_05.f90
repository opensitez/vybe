! vybe-test: fortran/expressions/expr_pow_05
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a=2,b
b = a ** 3
if ((b) /= 8) then
    print *, "FAIL: want [8] got [", b, "]"
    stop 1
end if
end program p
