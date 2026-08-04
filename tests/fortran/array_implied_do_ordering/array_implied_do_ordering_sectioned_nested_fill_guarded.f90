! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_sectioned_nested_fill_guarded
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_sectioned_nested_fill_guarded
    integer :: values(6)
    integer :: n
    n = 3
    values = 0
    if (n == 3) values = [(i*i, i = 1,6)]
    if ((sum(values)) /= 91) then
    print *, "FAIL: want [91] got [", sum(values), "]"
    stop 1
end if
    if ((values(3)) /= 9) then
    print *, "FAIL: want [9] got [", values(3), "]"
    stop 1
end if
    if ((values(6)) /= 36) then
    print *, "FAIL: want [36] got [", values(6), "]"
    stop 1
end if
end program array_implied_do_ordering_sectioned_nested_fill_guarded
