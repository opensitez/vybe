! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_descending_fill
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_descending_fill
    integer :: values(4)
    values = [(i, i=8,2,-2)]
    if ((size(values)) /= 4) then
    print *, "FAIL: want [4] got [", size(values), "]"
    stop 1
end if
    if ((values(1)) /= 8) then
    print *, "FAIL: want [8] got [", values(1), "]"
    stop 1
end if
    if ((values(4)) /= 2) then
    print *, "FAIL: want [2] got [", values(4), "]"
    stop 1
end if
    if ((sum(values)) /= 20) then
    print *, "FAIL: want [20] got [", sum(values), "]"
    stop 1
end if
end program array_implied_do_ordering_descending_fill
