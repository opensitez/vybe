! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_mixed_repeats_and_implied_do_tail
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program t
    integer, allocatable :: values(:)
    values = (/ (2, i = 1, 3), (i, i = 1, 3), (12, i = 1, 1) /)
    if ((size(values)) /= 7) then
    print *, "FAIL: want [7] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 24) then
    print *, "FAIL: want [24] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 12) then
    print *, "FAIL: want [12] got [", values(size(values)), "]"
    stop 1
end if
    if ((values(4)) /= 1) then
    print *, "FAIL: want [1] got [", values(4), "]"
    stop 1
end if
    if ((values(5)) /= 2) then
    print *, "FAIL: want [2] got [", values(5), "]"
    stop 1
end if
end program t
