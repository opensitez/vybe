! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_repeat_and_division_tail
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_repeat_and_division_tail
    integer, allocatable :: values(:)
    values = (/ 2 * 12, 1 * (20 / 4), 2 * 2 /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 34) then
    print *, "FAIL: want [34] got [", sum(values), "]"
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
