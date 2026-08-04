! vybe-test: fortran/else_where_fallback_order/test_else_where_single_match_applies_before_no_default_mask
! origin: languages/fortran/tests/fortran/test_else_where_fallback_order.rs

program test_else_where_single_match_applies
    integer :: a(3)
    integer :: b(3)
    a = (/1, 5, 9/)
    b = (/7, 8, 9/)
    where (a == 5)
        b = 50
    elsewhere (a == 6)
        b = 60
    end where
    if ((b(1)) /= 7) then
    print *, "FAIL: want [7] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 50) then
    print *, "FAIL: want [50] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 9) then
    print *, "FAIL: want [9] got [", b(3), "]"
    stop 1
end if
end program test_else_where_single_match_applies
