! vybe-test: fortran/expressions/expr_implied_do_29
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a(3)
a = [(i, i=1,3)]
if ((a(1) + a(2) + a(3)) /= 6) then
    print *, "FAIL: want [6] got [", a(1) + a(2) + a(3), "]"
    stop 1
end if
end program p
