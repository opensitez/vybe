! vybe-test: fortran/static_initialization_order/static_initialization_order_multiple_saves
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

program static_initialization_order_multiple_saves
    integer, save :: a = 1
    integer, save :: b = a + 1
    integer, save :: c = b + 1
    if ((a) /= 1) then
    print *, "FAIL: want [1] got [", a, "]"
    stop 1
end if
    if ((b) /= 2) then
    print *, "FAIL: want [2] got [", b, "]"
    stop 1
end if
    if ((c) /= 3) then
    print *, "FAIL: want [3] got [", c, "]"
    stop 1
end if
end program static_initialization_order_multiple_saves
