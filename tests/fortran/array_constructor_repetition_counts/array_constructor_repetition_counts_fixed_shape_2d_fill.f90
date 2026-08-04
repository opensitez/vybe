! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_fixed_shape_2d_fill
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_fixed_shape_2d_fill
    integer :: values(2, 3)
    values = (/ 1 * 5, 3 * 2, 2 * 1 /)
    if ((values(1, 1)) /= 5) then
    print *, "FAIL: want [5] got [", values(1, 1), "]"
    stop 1
end if
    if ((values(2, 1)) /= 2) then
    print *, "FAIL: want [2] got [", values(2, 1), "]"
    stop 1
end if
    if ((values(2, 3)) /= 1) then
    print *, "FAIL: want [1] got [", values(2, 3), "]"
    stop 1
end if
    if ((sum(values)) /= 13) then
    print *, "FAIL: want [13] got [", sum(values), "]"
    stop 1
end if
end program array_constructor_repetition_counts_fixed_shape_2d_fill
