! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_repeated_zero_length_while_not_empty
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program t
    integer, allocatable :: values(:)
    values = (/ (0, i = 1, 4), (9, i = 1, 2) /)
    if ((size(values)) /= 6) then
    print *, "FAIL: want [6] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 18) then
    print *, "FAIL: want [18] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 0) then
    print *, "FAIL: want [0] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 9) then
    print *, "FAIL: want [9] got [", values(size(values)), "]"
    stop 1
end if
end program t
