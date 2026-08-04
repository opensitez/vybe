! vybe-test: fortran/static_initialization_order/static_initialization_order_dependency_chain
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

program static_initialization_order_dependency_chain
    integer :: a = 1
    integer :: b = a + 2
    integer :: c = b * 3
    if ((a) /= 1) then
    print *, "FAIL: want [1] got [", a, "]"
    stop 1
end if
    if ((b) /= 3) then
    print *, "FAIL: want [3] got [", b, "]"
    stop 1
end if
    if ((c) /= 9) then
    print *, "FAIL: want [9] got [", c, "]"
    stop 1
end if
end program static_initialization_order_dependency_chain
