! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_complex_sequence
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_complex_sequence
    integer :: values(5)
    values = [(i*3 + i/2, i = 1, 5)]
    if ((values(1)) /= 3) then
    print *, "FAIL: want [3] got [", values(1), "]"
    stop 1
end if
    if ((values(5)) /= 17) then
    print *, "FAIL: want [17] got [", values(5), "]"
    stop 1
end if
    if ((sum(values)) /= 51) then
    print *, "FAIL: want [51] got [", sum(values), "]"
    stop 1
end if
end program array_implied_do_ordering_complex_sequence
