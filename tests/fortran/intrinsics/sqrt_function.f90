! vybe-test: fortran/intrinsics/sqrt_function
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    if ((sqrt(25.0)) /= 5) then
    print *, "FAIL: want [5] got [", sqrt(25.0), "]"
    stop 1
end if
end program test
