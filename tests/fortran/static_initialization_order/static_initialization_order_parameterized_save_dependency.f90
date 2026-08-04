! vybe-test: fortran/static_initialization_order/static_initialization_order_parameterized_save_dependency
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

module static_init_param_dependency
    integer, parameter :: scale = 2
    integer, save :: base = scale
    integer, save :: doubled = base * scale
end module static_init_param_dependency

program static_initialization_order_parameterized_save_dependency
    use static_init_param_dependency
    if ((base) /= 2) then
    print *, "FAIL: want [2] got [", base, "]"
    stop 1
end if
    if ((doubled) /= 4) then
    print *, "FAIL: want [4] got [", doubled, "]"
    stop 1
end if
end program static_initialization_order_parameterized_save_dependency
