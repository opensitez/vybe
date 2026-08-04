! vybe-test: fortran/arithmetic/assign_expression
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
integer :: x
x = 2 + 3 * 4
if ((x) /= 14) then
    print *, "FAIL: want [14] got [", x, "]"
    stop 1
end if
end program t
