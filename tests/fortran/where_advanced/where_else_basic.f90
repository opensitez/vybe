! vybe-test: fortran/where_advanced/where_else_basic
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    where (a > 3)
        b = a * 10
    elsewhere
        b = a
    end where
    if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
    if ((b(5)) /= 50) then
    print *, "FAIL: want [50] got [", b(5), "]"
    stop 1
end if
end program test
