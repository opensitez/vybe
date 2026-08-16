! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_implied_do_and_nested_repeats
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program t
    integer, allocatable :: values(:)
    values = (/ (i, i = 2, 4), (10, i = 1, 2), (3 + 2, i = 1, 1), (0, i = 1, 1) /)
    if ((size(values)) /= 7) then
    print *, "FAIL: want [7] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 34) then
    print *, "FAIL: want [34] got [", sum(values), "]"
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
end program t
