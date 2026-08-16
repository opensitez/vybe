! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_variable_expression_counts
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_variable_expression_counts
    integer, allocatable :: values(:)
    values = (/ (4, i = 1, 2 + 1), (1, i = 1, 1 + 2), (-3, i = 1, 2 + 1) /)
    if ((size(values)) /= 9) then
    print *, "FAIL: want [9] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 6) then
    print *, "FAIL: want [6] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 4) then
    print *, "FAIL: want [4] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= -3) then
    print *, "FAIL: want [-3] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_variable_expression_counts
