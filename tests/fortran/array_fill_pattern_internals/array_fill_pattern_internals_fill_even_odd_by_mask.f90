! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_even_odd_by_mask
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_even_odd_by_mask
    integer, allocatable :: values(:)
    integer :: i
    values = (/ 1, 2, 3, 4, 5, 6 /)
    where (mod(values, 2) == 0)
        values = 20
    elsewhere
        values = -20
    end where
    if ((sum(values)) /= 0) then
    print *, "FAIL: want [0] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= -20) then
    print *, "FAIL: want [-20] got [", values(1), "]"
    stop 1
end if
    if ((values(2)) /= 20) then
    print *, "FAIL: want [20] got [", values(2), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_even_odd_by_mask
