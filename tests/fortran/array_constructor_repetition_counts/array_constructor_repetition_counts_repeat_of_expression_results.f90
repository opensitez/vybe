! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_repeat_of_expression_results
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program t
    integer, allocatable :: values(:)
    values = (/ (1 + 2, i = 1, 3), (5 - 3, i = 1, 2) /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 13) then
    print *, "FAIL: want [13] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 3) then
    print *, "FAIL: want [3] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 2) then
    print *, "FAIL: want [2] got [", values(size(values)), "]"
    stop 1
end if
end program t
