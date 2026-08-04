! vybe-test: fortran/where_advanced/where_mask_with_abs
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    real :: a(6) = [-3., 2., -1., 4., -5., 0.]
    real :: b(6)
    b = 0.0
where (abs(a) > 2.0)
        b = a
    end where
    if ((nint(b(1))) /= -3) then
    print *, "FAIL: want [-3] got [", nint(b(1)), "]"
    stop 1
end if
    if ((nint(b(2))) /= 0) then
    print *, "FAIL: want [0] got [", nint(b(2)), "]"
    stop 1
end if
end program test
