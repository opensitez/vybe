! vybe-test: fortran/static_initialization_order/static_initialization_order_order_inside_block
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

program static_initialization_order_order_inside_block
    integer :: base = 10
    block
        integer :: nested = base + 1
        if ((nested) /= 11) then
    print *, "FAIL: want [11] got [", nested, "]"
    stop 1
end if
    end block
    if ((base) /= 10) then
    print *, "FAIL: want [10] got [", base, "]"
    stop 1
end if
end program static_initialization_order_order_inside_block
