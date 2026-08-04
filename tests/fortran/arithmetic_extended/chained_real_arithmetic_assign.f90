! vybe-test: fortran/arithmetic_extended/chained_real_arithmetic_assign
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
real :: y
y = 1.5 * 2.0 + 3.5
if (abs((y) - 6.5) > 1.0e-6) then
    print *, "FAIL: want [6.5] got [", y, "]"
    stop 1
end if
end program t
