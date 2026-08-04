! vybe-test: fortran/intrinsics_extended/abs_real
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
real :: x
x = abs(-3.5)
if (abs((x) - 3.5) > 1.0e-6) then
    print *, "FAIL: want [3.5] got [", x, "]"
    stop 1
end if
end program t
