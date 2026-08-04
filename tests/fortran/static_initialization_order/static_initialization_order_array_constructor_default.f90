! vybe-test: fortran/static_initialization_order/static_initialization_order_array_constructor_default
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

program static_initialization_order_array_constructor_default
    integer, save :: values(4) = (/1, 2, 3, 4/)
    integer :: i
    i = values(1) + values(4)
    if ((values(2)) /= 2) then
    print *, "FAIL: want [2] got [", values(2), "]"
    stop 1
end if
    if ((i) /= 5) then
    print *, "FAIL: want [5] got [", i, "]"
    stop 1
end if
end program static_initialization_order_array_constructor_default
