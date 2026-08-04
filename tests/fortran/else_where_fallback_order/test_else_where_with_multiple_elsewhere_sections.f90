! vybe-test: fortran/else_where_fallback_order/test_else_where_with_multiple_elsewhere_sections
! origin: languages/fortran/tests/fortran/test_else_where_fallback_order.rs

program test_else_where_fallback_multiple_sections
    integer :: a(5)
    integer :: b(5)
    a = (/1, 2, 3, 4, 5/)
    b = 0
    where (a <= 2)
        b = 5
    elsewhere where (a > 3)
        b = 10
    elsewhere where (a == 3)
        b = 7
    elsewhere
        b = 99
    end where
    if ((b(1)) /= 5) then
    print *, "FAIL: want [5] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 5) then
    print *, "FAIL: want [5] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 7) then
    print *, "FAIL: want [7] got [", b(3), "]"
    stop 1
end if
    if ((b(4)) /= 10) then
    print *, "FAIL: want [10] got [", b(4), "]"
    stop 1
end if
    if ((b(5)) /= 10) then
    print *, "FAIL: want [10] got [", b(5), "]"
    stop 1
end if
end program test_else_where_fallback_multiple_sections
