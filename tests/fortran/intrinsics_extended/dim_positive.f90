! vybe-test: fortran/intrinsics_extended/dim_positive
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((dim(10, 3)) /= 7) then
    print *, "FAIL: want [7] got [", dim(10, 3), "]"
    stop 1
end if
end program t
