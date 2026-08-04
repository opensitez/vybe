! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_reverse_then_forward
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_reverse_then_forward
    integer :: values(4)
    values = [(i, i=4,1,-1)]
    if ((values(1)) /= 4) then
    print *, "FAIL: want [4] got [", values(1), "]"
    stop 1
end if
    if ((values(4)) /= 1) then
    print *, "FAIL: want [1] got [", values(4), "]"
    stop 1
end if
    if ((sum(values)) /= 10) then
    print *, "FAIL: want [10] got [", sum(values), "]"
    stop 1
end if
end program array_implied_do_ordering_reverse_then_forward
