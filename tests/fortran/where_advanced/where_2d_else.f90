! vybe-test: fortran/where_advanced/where_2d_else
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    real :: m(4,4) = reshape([(real(i), i=1,16)],[4,4])
    real :: result(4,4)
where (m > 8.0)
        result = m
    elsewhere
        result = 0.0
    end where
    if ((nint(result(1,1))) /= 0) then
    print *, "FAIL: want [0] got [", nint(result(1,1)), "]"
    stop 1
end if
    if ((nint(result(4,4))) /= 16) then
    print *, "FAIL: want [16] got [", nint(result(4,4)), "]"
    stop 1
end if
end program test
