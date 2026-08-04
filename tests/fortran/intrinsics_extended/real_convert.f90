! vybe-test: fortran/intrinsics_extended/real_convert
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((real(7)) /= 7) then
    print *, "FAIL: want [7] got [", real(7), "]"
    stop 1
end if
end program t
