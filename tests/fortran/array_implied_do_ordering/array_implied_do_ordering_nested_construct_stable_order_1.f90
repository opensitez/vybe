! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_nested_construct_stable_order_1
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_nested_construct_stable_order_1
    integer :: values(6)
    integer :: i
    i = size([(j, j = 1, 6)])
    values = [(i, i = 1, 6)]
    if ((i) /= 6) then
    print *, "FAIL: want [6] got [", i, "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(6)) /= 6) then
    print *, "FAIL: want [6] got [", values(6), "]"
    stop 1
end if
    if ((sum(values)) /= 21) then
    print *, "FAIL: want [21] got [", sum(values), "]"
    stop 1
end if
end program array_implied_do_ordering_nested_construct_stable_order_1
