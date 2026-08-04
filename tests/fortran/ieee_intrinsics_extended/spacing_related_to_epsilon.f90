! vybe-test: fortran/ieee_intrinsics_extended/spacing_related_to_epsilon
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((merge(1, 0, spacing(1.0) <= epsilon(1.0))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, spacing(1.0) <= epsilon(1.0)), "]"
    stop 1
end if
end program t
