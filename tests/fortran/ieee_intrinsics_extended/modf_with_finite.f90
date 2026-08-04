! vybe-test: fortran/ieee_intrinsics_extended/modf_with_finite
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
real :: f
integer :: i
f = modf(3.75, i)
if ((i) /= 3) then
    print *, "FAIL: want [3] got [", i, "]"
    stop 1
end if
if ((merge(1, 0, ieee_is_finite(f))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, ieee_is_finite(f)), "]"
    stop 1
end if
end program t
