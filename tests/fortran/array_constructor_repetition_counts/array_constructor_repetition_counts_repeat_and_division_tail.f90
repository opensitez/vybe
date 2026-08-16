! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_repeat_and_division_tail
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_repeat_and_division_tail
    integer, allocatable :: values(:)
    values = (/ (12, i = 1, 2), (20 / 4, i = 1, 1), (2, i = 1, 2) /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 33) then
    print *, "FAIL: want [33] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 12) then
    print *, "FAIL: want [12] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 2) then
    print *, "FAIL: want [2] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_repeat_and_division_tail
