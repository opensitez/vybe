! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_multi_expression_series
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_multi_expression_series
    integer :: values(3)
    values = [(i*i + 1, i = 1, 3)]
    if ((values(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(2)) /= 5) then
    print *, "FAIL: want [5] got [", values(2), "]"
    stop 1
end if
    if ((values(3)) /= 10) then
    print *, "FAIL: want [10] got [", values(3), "]"
    stop 1
end if
    if ((sum(values)) /= 17) then
    print *, "FAIL: want [17] got [", sum(values), "]"
    stop 1
end if
end program array_implied_do_ordering_multi_expression_series
