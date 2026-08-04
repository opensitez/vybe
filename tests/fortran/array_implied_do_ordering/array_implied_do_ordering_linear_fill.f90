! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_linear_fill
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_linear_fill
    integer :: values(5)
    values = [(i, i=1,5)]
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(5)) /= 5) then
    print *, "FAIL: want [5] got [", values(5), "]"
    stop 1
end if
    if ((sum(values)) /= 15) then
    print *, "FAIL: want [15] got [", sum(values), "]"
    stop 1
end if
end program array_implied_do_ordering_linear_fill
