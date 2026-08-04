! vybe-test: fortran/else_where_fallback_order/test_else_where_default_only_runs_when_previous_masks_fail
! origin: languages/fortran/tests/fortran/test_else_where_fallback_order.rs

program test_else_where_fallback_default_only
    integer :: a(4)
    integer :: b(4)
    a = (/0, 1, 2, 3/)
    b = 0
    where (a == 10)
        b = 1
    elsewhere where (a == 20)
        b = 2
    elsewhere
        b = 9
    end where
    if ((b(1)) /= 9) then
    print *, "FAIL: want [9] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 9) then
    print *, "FAIL: want [9] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 9) then
    print *, "FAIL: want [9] got [", b(3), "]"
    stop 1
end if
    if ((b(4)) /= 9) then
    print *, "FAIL: want [9] got [", b(4), "]"
    stop 1
end if
end program test_else_where_fallback_default_only
