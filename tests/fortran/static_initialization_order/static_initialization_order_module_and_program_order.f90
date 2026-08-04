! vybe-test: fortran/static_initialization_order/static_initialization_order_module_and_program_order
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

module m
    integer, parameter :: base = 3
    integer, save :: offset = base + 1
    integer, save :: total = offset * 4
end module m

program static_initialization_order_module_and_program_order
    use m
    if ((total) /= 16) then
    print *, "FAIL: want [16] got [", total, "]"
    stop 1
end if
    if ((base) /= 3) then
    print *, "FAIL: want [3] got [", base, "]"
    stop 1
end if
    if ((offset) /= 4) then
    print *, "FAIL: want [4] got [", offset, "]"
    stop 1
end if
end program static_initialization_order_module_and_program_order
