! vybe-test: fortran/intrinsics_extended/log10_100
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
real :: x
x = log10(100.0)
if ((x) /= 2) then
    print *, "FAIL: want [2] got [", x, "]"
    stop 1
end if
end program t
