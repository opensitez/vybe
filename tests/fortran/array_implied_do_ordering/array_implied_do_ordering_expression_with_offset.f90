! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_expression_with_offset
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_expression_with_offset
    integer :: values(5)
    values = [(i + 3, i = 0, 4)]
    if ((values(1)) /= 3) then
    print *, "FAIL: want [3] got [", values(1), "]"
    stop 1
end if
    if ((values(3)) /= 5) then
    print *, "FAIL: want [5] got [", values(3), "]"
    stop 1
end if
    if ((values(5)) /= 7) then
    print *, "FAIL: want [7] got [", values(5), "]"
    stop 1
end if
    if ((sum(values)) /= 25) then
    print *, "FAIL: want [25] got [", sum(values), "]"
    stop 1
end if
end program array_implied_do_ordering_expression_with_offset
