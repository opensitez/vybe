! vybe-test: fortran/intrinsics_extended/ibset_bit
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((ibset(0, 3)) /= 8) then
    print *, "FAIL: want [8] got [", ibset(0, 3), "]"
    stop 1
end if
end program t
