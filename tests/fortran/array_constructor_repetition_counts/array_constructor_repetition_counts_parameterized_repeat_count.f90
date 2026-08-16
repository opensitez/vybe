! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_parameterized_repeat_count
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_parameterized_repeat_count
    integer, parameter :: n = 4
    integer, allocatable :: values(:)
    values = (/ (3, i = 1, n), (1, i = 1, 2) /)
    if ((size(values)) /= 6) then
    print *, "FAIL: want [6] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 14) then
    print *, "FAIL: want [14] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 3) then
    print *, "FAIL: want [3] got [", values(1), "]"
    stop 1
end if
    if ((values(n + 1)) /= 1) then
    print *, "FAIL: want [1] got [", values(n + 1), "]"
    stop 1
end if
end program array_constructor_repetition_counts_parameterized_repeat_count
