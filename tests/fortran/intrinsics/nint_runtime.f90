! vybe-test: fortran/intrinsics/nint_runtime
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    if ((nint(2.6)) /= 3) then
    print *, "FAIL: want [3] got [", nint(2.6), "]"
    stop 1
end if
end program test
