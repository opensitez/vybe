! vybe-test: fortran/else_where_fallback_order/test_else_where_masked_even_odd_fallthrough
! origin: languages/fortran/tests/fortran/test_else_where_fallback_order.rs

program test_else_where_masked_even_odd_fallthrough
    integer :: a(5)
    integer :: b(5)
    a = (/1, 2, 3, 4, 5/)
    b = 0
    where (a > 3)
        b = 10
    elsewhere (mod(a, 2) == 1)
        b = 20
    elsewhere
        b = 30
    end where
    if ((b(1)) /= 20) then
    print *, "FAIL: want [20] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 30) then
    print *, "FAIL: want [30] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 20) then
    print *, "FAIL: want [20] got [", b(3), "]"
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
end program test_else_where_masked_even_odd_fallthrough
