! vybe-test: fortran/expressions/expr_section_22
! origin: languages/fortran/tests/fortran/test_expressions.rs
program p
integer :: a(4)
a = [1,2,3,4]
if ((a(2) + a(3)) /= 5) then
    print *, "FAIL: want [5] got [", a(2) + a(3), "]"
    stop 1
end if
end program p
