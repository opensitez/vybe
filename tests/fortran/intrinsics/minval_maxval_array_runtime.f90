! vybe-test: fortran/intrinsics/minval_maxval_array_runtime
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    integer :: a(4) = [7, 3, 9, 4]
    if ((minval(a)) /= 3) then
    print *, "FAIL: want [3] got [", minval(a), "]"
    stop 1
end if
    if ((maxval(a)) /= 9) then
    print *, "FAIL: want [9] got [", maxval(a), "]"
    stop 1
end if
end program test
