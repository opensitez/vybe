! vybe-test: fortran/expressions/expr_index_23
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a(3)
a = [1,2,3]
if ((a(2)) /= 2) then
    print *, "FAIL: want [2] got [", a(2), "]"
    stop 1
end if
end program p
