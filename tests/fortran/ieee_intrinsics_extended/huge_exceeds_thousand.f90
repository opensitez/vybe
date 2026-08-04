! vybe-test: fortran/ieee_intrinsics_extended/huge_exceeds_thousand
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((merge(1, 0, huge(1.0) > 1000.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, huge(1.0) > 1000.0), "]"
    stop 1
end if
end program t
