! vybe-test: fortran/expressions/expr_power_assoc_34
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a
a = 2 ** 3 ** 2
if ((a) /= 512) then
    print *, "FAIL: want [512] got [", a, "]"
    stop 1
end if
end program p
