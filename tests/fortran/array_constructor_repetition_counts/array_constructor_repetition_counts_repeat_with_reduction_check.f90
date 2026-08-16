! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_repeat_with_reduction_check
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_repeat_with_reduction_check
    integer, allocatable :: values(:)
    values = (/ (1, i = 1, 5), (2, i = 1, 3), (3, i = 1, 2) /)
    if ((size(values)) /= 10) then
    print *, "FAIL: want [10] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 17) then
    print *, "FAIL: want [17] got [", sum(values), "]"
    stop 1
end if
    if ((count(values == 2)) /= 3) then
    print *, "FAIL: want [3] got [", count(values == 2), "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
end program array_constructor_repetition_counts_repeat_with_reduction_check
