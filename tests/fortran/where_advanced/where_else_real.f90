! vybe-test: fortran/where_advanced/where_else_real
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    real :: x(6) = [-2., -1., 0., 1., 2., 3.]
    real :: y(6)
where (x >= 0.0)
        y = sqrt(x)
    elsewhere
        y = 0.0
    end where
    if ((nint(y(1))) /= 0) then
    print *, "FAIL: want [0] got [", nint(y(1)), "]"
    stop 1
end if
    if ((nint(y(4))) /= 1) then
    print *, "FAIL: want [1] got [", nint(y(4)), "]"
    stop 1
end if
end program test
