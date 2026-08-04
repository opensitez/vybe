! vybe-test: fortran/intrinsics/abs_function
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    if ((abs(-42)) /= 42) then
    print *, "FAIL: want [42] got [", abs(-42), "]"
    stop 1
end if
end program test
