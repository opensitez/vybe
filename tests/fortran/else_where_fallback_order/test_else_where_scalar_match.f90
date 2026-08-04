! vybe-test: fortran/else_where_fallback_order/test_else_where_scalar_match
! origin: languages/fortran/tests/fortran/test_else_where_fallback_order.rs

program test_else_where_scalar_match
    integer :: a(1)
    integer :: b(1)
    a = (/7/)
    where (a == 7)
        b = 77
    elsewhere
        b = 13
    end where
    if ((b(1)) /= 77) then
    print *, "FAIL: want [77] got [", b(1), "]"
    stop 1
end if
end program test_else_where_scalar_match
