! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_nested_index_expression
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_nested_index_expression
    integer :: values(5)
    values = [(i*2, i = 1,5)]
    if ((values(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(5)) /= 10) then
    print *, "FAIL: want [10] got [", values(5), "]"
    stop 1
end if
    if ((sum(values)) /= 30) then
    print *, "FAIL: want [30] got [", sum(values), "]"
    stop 1
end if
end program array_implied_do_ordering_nested_index_expression
