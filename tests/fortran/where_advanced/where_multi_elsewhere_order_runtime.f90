! vybe-test: fortran/where_advanced/where_multi_elsewhere_order_runtime
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(6) = [1, 5, 10, 50, 100, 500]
    integer :: b(6)
    where (a < 10)
        b = 1
    elsewhere (a == 100)
        b = 2
    elsewhere
        b = 3
    end where
    if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
    if ((b(3)) /= 1) then
    print *, "FAIL: want [1] got [", b(3), "]"
    stop 1
end if
    if ((b(5)) /= 3) then
    print *, "FAIL: want [3] got [", b(5), "]"
    stop 1
end if
    if ((b(6)) /= 3) then
    print *, "FAIL: want [3] got [", b(6), "]"
    stop 1
end if
end program test
