! vybe-test: fortran/intrinsics_extended/cosh_zero
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
real :: x
x = cosh(0.0)
if ((x) /= 1) then
    print *, "FAIL: want [1] got [", x, "]"
    stop 1
end if
end program t
