! vybe-test: fortran/intrinsics_extended/aint_truncate
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
real :: x
x = aint(3.9)
if ((x) /= 3) then
    print *, "FAIL: want [3] got [", x, "]"
    stop 1
end if
end program t
