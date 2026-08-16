! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_fill_via_function_like_expression
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_fill_via_function_like_expression
    integer :: values(4)
    values = [(i + 10, i = 1,4)]
    if ((sum(values)) /= 50) then
    print *, "FAIL: want [50] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 11) then
    print *, "FAIL: want [11] got [", values(1), "]"
    stop 1
end if
    if ((values(4)) /= 14) then
    print *, "FAIL: want [14] got [", values(4), "]"
    stop 1
end if
end program array_implied_do_ordering_fill_via_function_like_expression
