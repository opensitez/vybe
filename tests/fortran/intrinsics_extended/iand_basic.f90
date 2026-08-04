! vybe-test: fortran/intrinsics_extended/iand_basic
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((iand(255, 15)) /= 15) then
    print *, "FAIL: want [15] got [", iand(255, 15), "]"
    stop 1
end if
end program t
