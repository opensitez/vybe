! vybe-test: fortran/else_where_fallback_order/test_else_where_no_default_no_match_preserves_existing_values
! origin: languages/fortran/tests/fortran/test_else_where_fallback_order.rs

program test_else_where_no_default_no_match_preserves
    integer :: a(4)
    integer :: b(4)
    a = (/1, 2, 3, 4/)
    b = (/10, 20, 30, 40/)
    where (a > 4)
        b = 10 * a
    elsewhere (a < 0)
        b = -10
    end where
    if ((b(1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 20) then
    print *, "FAIL: want [20] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 30) then
    print *, "FAIL: want [30] got [", b(3), "]"
    stop 1
end if
    if ((b(4)) /= 40) then
    print *, "FAIL: want [40] got [", b(4), "]"
    stop 1
end if
end program test_else_where_no_default_no_match_preserves
