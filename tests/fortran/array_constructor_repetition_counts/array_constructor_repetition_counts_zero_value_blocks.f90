! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_zero_value_blocks
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_zero_value_blocks
    integer, allocatable :: values(:)
    values = (/ (0, i = 1, 3), (7, i = 1, 2), (0, i = 1, 1) /)
    if ((size(values)) /= 6) then
    print *, "FAIL: want [6] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 14) then
    print *, "FAIL: want [14] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 0) then
    print *, "FAIL: want [0] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 0) then
    print *, "FAIL: want [0] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_zero_value_blocks
