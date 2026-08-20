! vybe-test: fortran/else_where_fallback_order/test_else_where_fallback_order_applies_primary_mask
! origin: languages/fortran/tests/fortran/test_else_where_fallback_order.rs

program test_else_where_fallback_order
    integer :: a(4)
    integer :: b(4)
    a = (/1, 2, 3, 4/)
    b = 0
    where (a > 2)
        b = 1
    elsewhere (a == 2)
        b = 2
    elsewhere
        b = 3
    end where
    if ((b(1)) /= 3) then
    print *, "FAIL: want [3] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 2) then
    print *, "FAIL: want [2] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 1) then
    print *, "FAIL: want [1] got [", b(3), "]"
    stop 1
end if
    if ((b(4)) /= 1) then
    print *, "FAIL: want [1] got [", b(4), "]"
    stop 1
end if
end program test_else_where_fallback_order
