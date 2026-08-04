! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_repeat_with_sign
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_repeat_with_sign
    integer, allocatable :: values(:)
    values = (/ 4 * (-1), 1 * 6, 1 * (-3) /)
    if ((size(values)) /= 6) then
    print *, "FAIL: want [6] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 3) then
    print *, "FAIL: want [3] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= -1) then
    print *, "FAIL: want [-1] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= -3) then
    print *, "FAIL: want [-3] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_repeat_with_sign
