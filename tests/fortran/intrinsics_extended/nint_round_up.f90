! vybe-test: fortran/intrinsics_extended/nint_round_up
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((nint(3.7)) /= 4) then
    print *, "FAIL: want [4] got [", nint(3.7), "]"
    stop 1
end if
end program t
