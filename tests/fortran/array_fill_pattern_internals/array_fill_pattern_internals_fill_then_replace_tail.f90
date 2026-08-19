! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_then_replace_tail
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_then_replace_tail
    integer, allocatable :: values(:)
    values = (/ 1, 1, 1, 1, 1, 1 /)
    values(1:3) = -3
    values(4:6) = values(1:3) + 5
    if ((sum(values)) /= -3) then
    print *, "FAIL: want [-3] got [", sum(values), "]"
    stop 1
end if
    if ((values(3)) /= -3) then
    print *, "FAIL: want [-3] got [", values(3), "]"
    stop 1
end if
    if ((values(6)) /= 2) then
    print *, "FAIL: want [2] got [", values(6), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_then_replace_tail
