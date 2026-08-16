! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_and_reset_with_where
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_and_reset_with_where
    integer, allocatable :: values(:)
    values = (/ 9, 8, 7, 6, 5, 4 /)
    where (values > 6)
        values = 0
    end where
    if ((sum(values)) /= 15) then
    print *, "FAIL: want [15] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 0) then
    print *, "FAIL: want [0] got [", values(1), "]"
    stop 1
end if
    if ((values(2)) /= 0) then
    print *, "FAIL: want [0] got [", values(2), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_and_reset_with_where
