! vybe-test: fortran/intrinsics_extended/nint_round_down
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((nint(3.2)) /= 3) then
    print *, "FAIL: want [3] got [", nint(3.2), "]"
    stop 1
end if
end program t
