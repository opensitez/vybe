! vybe-test: fortran/else_where_fallback_order/test_else_where_chain_skips_unreached_sections
! origin: languages/fortran/tests/fortran/test_else_where_fallback_order.rs

program test_else_where_chain_skips_unreached
    integer :: a(5)
    integer :: b(5)
    a = (/1, 2, 3, 4, 5/)
    b = 0
    where (a > 10)
        b = 1
    elsewhere (a > 3)
        b = 2
    elsewhere (a == 4)
        b = 9
    elsewhere
        b = 3
    end where
    if ((b(1)) /= 3) then
    print *, "FAIL: want [3] got [", b(1), "]"
    stop 1
end if
    if ((b(4)) /= 2) then
    print *, "FAIL: want [2] got [", b(4), "]"
    stop 1
end if
    if ((b(5)) /= 3) then
    print *, "FAIL: want [3] got [", b(5), "]"
    stop 1
end if
end program test_else_where_chain_skips_unreached
