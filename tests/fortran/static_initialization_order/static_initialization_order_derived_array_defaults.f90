! vybe-test: fortran/static_initialization_order/static_initialization_order_derived_array_defaults
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

program static_initialization_order_derived_array_defaults
    integer :: table(3) = (/1, 2, 3/)
    integer :: total
    total = sum(table)
    if ((table(1)) /= 1) then
    print *, "FAIL: want [1] got [", table(1), "]"
    stop 1
end if
    if ((table(3)) /= 3) then
    print *, "FAIL: want [3] got [", table(3), "]"
    stop 1
end if
    if ((total) /= 6) then
    print *, "FAIL: want [6] got [", total, "]"
    stop 1
end if
end program static_initialization_order_derived_array_defaults
