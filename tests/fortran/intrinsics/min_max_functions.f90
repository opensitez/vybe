! vybe-test: fortran/intrinsics/min_max_functions
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    if ((min(3, 7)) /= 3) then
    print *, "FAIL: want [3] got [", min(3, 7), "]"
    stop 1
end if
    if ((max(3, 7)) /= 7) then
    print *, "FAIL: want [7] got [", max(3, 7), "]"
    stop 1
end if
end program test
