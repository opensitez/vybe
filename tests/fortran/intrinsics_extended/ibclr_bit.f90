! vybe-test: fortran/intrinsics_extended/ibclr_bit
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((ibclr(15, 0)) /= 14) then
    print *, "FAIL: want [14] got [", ibclr(15, 0), "]"
    stop 1
end if
end program t
