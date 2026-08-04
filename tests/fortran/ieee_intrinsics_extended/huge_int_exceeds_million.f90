! vybe-test: fortran/ieee_intrinsics_extended/huge_int_exceeds_million
! origin: languages/fortran/tests/fortran/test_ieee_intrinsics_extended.rs
program t
if ((merge(1, 0, huge(0) > 1000000)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, huge(0) > 1000000), "]"
    stop 1
end if
end program t
