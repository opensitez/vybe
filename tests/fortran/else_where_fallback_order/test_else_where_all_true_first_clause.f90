! vybe-test: fortran/else_where_fallback_order/test_else_where_all_true_first_clause
! origin: languages/fortran/tests/fortran/test_else_where_fallback_order.rs

program test_else_where_all_true_first_clause
    integer :: a(4)
    integer :: b(4)
    a = (/10, 20, 30, 40/)
    where (a > 0)
        b = 5
    elsewhere
        b = 9
    end where
    if ((b(1)) /= 5) then
    print *, "FAIL: want [5] got [", b(1), "]"
    stop 1
end if
    if ((b(4)) /= 5) then
    print *, "FAIL: want [5] got [", b(4), "]"
    stop 1
end if
end program test_else_where_all_true_first_clause
