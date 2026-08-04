! vybe-test: fortran/where_advanced/where_2d_basic
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: m(3,3) = reshape([1,2,3,4,5,6,7,8,9],[3,3])
    where (m > 5)
        m = m * 2
    end where
    if ((m(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,1), "]"
    stop 1
end if
    if ((m(3,3)) /= 18) then
    print *, "FAIL: want [18] got [", m(3,3), "]"
    stop 1
end if
end program test
