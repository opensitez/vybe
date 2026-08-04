! vybe-test: fortran/intrinsics_extended/max_three
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
if ((max(5, 3, 8)) /= 8) then
    print *, "FAIL: want [8] got [", max(5, 3, 8), "]"
    stop 1
end if
end program t
