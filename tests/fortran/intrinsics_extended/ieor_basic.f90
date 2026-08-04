! vybe-test: fortran/intrinsics_extended/ieor_basic
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((ieor(255, 15)) /= 240) then
    print *, "FAIL: want [240] got [", ieor(255, 15), "]"
    stop 1
end if
end program t
