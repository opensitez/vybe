! vybe-test: fortran/intrinsics_extended/atan_zero
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
real :: x
x = atan(0.0)
if ((x) /= 0) then
    print *, "FAIL: want [0] got [", x, "]"
    stop 1
end if
end program t
