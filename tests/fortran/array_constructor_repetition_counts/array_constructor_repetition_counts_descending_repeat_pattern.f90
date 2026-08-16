! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_descending_repeat_pattern
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_descending_repeat_pattern
    integer, allocatable :: values(:)
    values = (/ (9, i = 1, 3), (8, i = 1, 1), (7, i = 1, 2), (6, i = 1, 1), 1 /)
    if ((size(values)) /= 8) then
    print *, "FAIL: want [8] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 56) then
    print *, "FAIL: want [56] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 9) then
    print *, "FAIL: want [9] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 1) then
    print *, "FAIL: want [1] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_descending_repeat_pattern
