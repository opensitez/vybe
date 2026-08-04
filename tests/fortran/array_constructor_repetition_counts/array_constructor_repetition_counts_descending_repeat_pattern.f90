! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_descending_repeat_pattern
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_descending_repeat_pattern
    integer, allocatable :: values(:)
    values = (/ 3 * 9, 1 * 8, 2 * 7, 1 * 6, 1 /)
    if ((size(values)) /= 8) then
    print *, "FAIL: want [8] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 55) then
    print *, "FAIL: want [55] got [", sum(values), "]"
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
