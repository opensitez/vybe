! vybe-test: fortran/static_initialization_order/static_initialization_order_component_with_intent_like_defaults
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

program static_initialization_order_component_with_intent_like_defaults
    type state
        integer :: left = 1
        integer :: right = 2
        integer :: total
    end type state
    type(state) :: s
    s%total = s%left + s%right
    if ((s%left) /= 1) then
    print *, "FAIL: want [1] got [", s%left, "]"
    stop 1
end if
    if ((s%right) /= 2) then
    print *, "FAIL: want [2] got [", s%right, "]"
    stop 1
end if
    if ((s%total) /= 3) then
    print *, "FAIL: want [3] got [", s%total, "]"
    stop 1
end if
end program static_initialization_order_component_with_intent_like_defaults
