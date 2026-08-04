! vybe-test: fortran/intrinsics_extended/ior_basic
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((ior(240, 15)) /= 255) then
    print *, "FAIL: want [255] got [", ior(240, 15), "]"
    stop 1
end if
end program t
