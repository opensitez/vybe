! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_negative_repeated_term
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_negative_repeated_term
    integer, allocatable :: values(:)
    values = (/ (-4, i = 1, 3), (5, i = 1, 2) /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= -2) then
    print *, "FAIL: want [-2] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= -4) then
    print *, "FAIL: want [-4] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 5) then
    print *, "FAIL: want [5] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_negative_repeated_term
