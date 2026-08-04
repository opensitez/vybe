! vybe-test: fortran/arithmetic/unary_plus_keeps_sign
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
integer :: x
x = +7
if ((x) /= 7) then
    print *, "FAIL: want [7] got [", x, "]"
    stop 1
end if
end program t
