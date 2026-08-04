! vybe-test: fortran/expressions/expr_complex_add_20
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
complex :: a=(1.0,2.0), b=(3.0,4.0), c
c = a + b
if ((real(c)) /= 4) then
    print *, "FAIL: want [4] got [", real(c), "]"
    stop 1
end if
if ((aimag(c)) /= 6) then
    print *, "FAIL: want [6] got [", aimag(c), "]"
    stop 1
end if
end program p
