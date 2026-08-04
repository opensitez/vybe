! vybe-test: fortran/intrinsics/mod_function
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    if ((mod(17, 5)) /= 2) then
    print *, "FAIL: want [2] got [", mod(17, 5), "]"
    stop 1
end if
end program test
