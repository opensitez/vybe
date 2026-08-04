! vybe-test: fortran/expressions/expr_array_constructor_21
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a(3)
a = [1,2,3]
if ((a(1) + a(2) + a(3)) /= 6) then
    print *, "FAIL: want [6] got [", a(1) + a(2) + a(3), "]"
    stop 1
end if
end program p
