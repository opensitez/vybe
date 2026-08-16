! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_implied_do_prefix_then_repeats
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program t
    integer, allocatable :: values(:)
    values = (/ (i, i = 1, 4), (5, i = 1, 2) /)
    if ((size(values)) /= 6) then
    print *, "FAIL: want [6] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 20) then
    print *, "FAIL: want [20] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 5) then
    print *, "FAIL: want [5] got [", values(size(values)), "]"
    stop 1
end if
end program t
