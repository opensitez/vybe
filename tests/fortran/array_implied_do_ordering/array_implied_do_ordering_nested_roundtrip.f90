! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_nested_roundtrip
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_nested_roundtrip
    integer :: values(6)
    values = [(i+j, i=1,2, j=1,3)]
    if ((size(values)) /= 6) then
    print *, "FAIL: want [6] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 21) then
    print *, "FAIL: want [21] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(6)) /= 5) then
    print *, "FAIL: want [5] got [", values(6), "]"
    stop 1
end if
end program array_implied_do_ordering_nested_roundtrip
