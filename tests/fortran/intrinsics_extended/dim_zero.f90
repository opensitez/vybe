! vybe-test: fortran/intrinsics_extended/dim_zero
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((dim(3, 10)) /= 0) then
    print *, "FAIL: want [0] got [", dim(3, 10), "]"
    stop 1
end if
end program t
