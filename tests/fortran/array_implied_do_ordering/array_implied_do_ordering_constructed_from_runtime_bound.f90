! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_constructed_from_runtime_bound
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_constructed_from_runtime_bound
    integer :: n
    integer :: values(4)
    n = 1
    values = [(i*n, i=1,4)]
    if ((sum(values)) /= 10) then
    print *, "FAIL: want [10] got [", sum(values), "]"
    stop 1
end if
    if ((values(4)) /= 4) then
    print *, "FAIL: want [4] got [", values(4), "]"
    stop 1
end if
    if ((values(2)) /= 2) then
    print *, "FAIL: want [2] got [", values(2), "]"
    stop 1
end if
end program array_implied_do_ordering_constructed_from_runtime_bound
