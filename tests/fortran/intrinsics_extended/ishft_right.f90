! vybe-test: fortran/intrinsics_extended/ishft_right
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((ishft(256, -4)) /= 16) then
    print *, "FAIL: want [16] got [", ishft(256, -4), "]"
    stop 1
end if
end program t
