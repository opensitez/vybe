! vybe-test: fortran/arithmetic/unary_minus
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
integer :: x = 5
if ((-x) /= -5) then
    print *, "FAIL: want [-5] got [", -x, "]"
    stop 1
end if
end program t
