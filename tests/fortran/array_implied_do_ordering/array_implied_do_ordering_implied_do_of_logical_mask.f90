! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_implied_do_of_logical_mask
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_implied_do_of_logical_mask
    integer :: values(4)
    values = [(merge(1, 0, mod(i,2) == 0), i = 1, 4)]
    if ((sum(values)) /= 2) then
    print *, "FAIL: want [2] got [", sum(values), "]"
    stop 1
end if
    if ((values(2)) /= 1) then
    print *, "FAIL: want [1] got [", values(2), "]"
    stop 1
end if
    if ((values(4)) /= 1) then
    print *, "FAIL: want [1] got [", values(4), "]"
    stop 1
end if
end program array_implied_do_ordering_implied_do_of_logical_mask
