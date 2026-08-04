! vybe-test: fortran/where_advanced/nested_where
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(6) = [1, 10, 2, 20, 3, 30]
    integer :: b(6) = 0
    where (a > 5)
        where (a > 15)
            b = a * 100
        elsewhere
            b = a * 10
        end where
    elsewhere
        b = a
    end where
    if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 100) then
    print *, "FAIL: want [100] got [", b(2), "]"
    stop 1
end if
    if ((b(4)) /= 2000) then
    print *, "FAIL: want [2000] got [", b(4), "]"
    stop 1
end if
end program test
