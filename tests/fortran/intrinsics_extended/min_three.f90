! vybe-test: fortran/intrinsics_extended/min_three
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((min(5, 3, 8)) /= 3) then
    print *, "FAIL: want [3] got [", min(5, 3, 8), "]"
    stop 1
end if
end program t
