! vybe-test: fortran/intrinsics_extended/int_convert
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((int(3.9)) /= 3) then
    print *, "FAIL: want [3] got [", int(3.9), "]"
    stop 1
end if
end program t
