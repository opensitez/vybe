! vybe-test: fortran/intrinsics_extended/ishft_left
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((ishft(1, 4)) /= 16) then
    print *, "FAIL: want [16] got [", ishft(1, 4), "]"
    stop 1
end if
end program t
