! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_implied_do_and_nested_repeats
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_implied_do_and_nested_repeats
    integer, allocatable :: values(:)
    values = (/ (i, i = 2, 4), 2 * 10, 1 * (3 + 2), 1 * 0 /)
    if ((size(values)) /= 8) then
    print *, "FAIL: want [8] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 33) then
    print *, "FAIL: want [33] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(6)) /= 0) then
    print *, "FAIL: want [0] got [", values(6), "]"
    stop 1
end if
end program array_constructor_repetition_counts_implied_do_and_nested_repeats
